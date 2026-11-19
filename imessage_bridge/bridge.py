"""Isolated Messages reader. No network, writes, attachments, or logging."""
import json
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

DB = Path.home() / 'Library/Messages/chat.db'
CONFIG = Path.home() / '.config/beaman-imessage/allowlist.json'
EPOCH = datetime(2001, 1, 1, tzinfo=timezone.utc)


class BridgeError(Exception):
    pass


def _config():
    try:
        data = json.loads(CONFIG.read_text())
        if not isinstance(data, dict) or set(data) != {'threads'} or not isinstance(data['threads'], dict):
            raise ValueError()
        for guid, participants in data['threads'].items():
            if not isinstance(guid, str) or not guid or not isinstance(participants, list) or not participants or any(not isinstance(p, str) or not p for p in participants) or len(set(participants)) != len(participants):
                raise ValueError()
        return data['threads']
    except (OSError, ValueError, TypeError):
        raise BridgeError('Allowlist unavailable or invalid') from None


def _connect():
    try:
        # mode=ro retains WAL visibility; query_only provides a second guard.
        con = sqlite3.connect(DB.as_uri() + '?mode=ro', uri=True)
        con.execute('PRAGMA query_only=ON')
        return con
    except (sqlite3.Error, ValueError):
        raise BridgeError('Messages database unavailable') from None


def list_allowed_message_threads():
    return [{'thread_id': guid} for guid in sorted(_config())]


def read_recent_messages(thread_id, limit=10):
    return _read(thread_id, limit, None)


def search_allowed_messages(thread_id, query, limit=10):
    if not isinstance(query, str) or not 1 <= len(query) <= 200:
        raise BridgeError('Invalid request')
    return _read(thread_id, limit, query)


def _read(thread_id, limit, query):
    allowed = _config()
    if not isinstance(thread_id, str) or thread_id not in allowed:
        raise BridgeError('Thread not allowed')
    if type(limit) is not int or not 1 <= limit <= 50:
        raise BridgeError('Invalid request')
    con = _connect()
    try:
        with con:
            chats = con.execute('SELECT ROWID FROM chat WHERE guid = ?', (thread_id,)).fetchall()
            if len(chats) != 1:
                raise BridgeError('Thread unavailable')
            chat_id = chats[0][0]
            participants = con.execute('SELECT h.id FROM chat_handle_join j JOIN handle h ON h.ROWID=j.handle_id WHERE j.chat_id=?', (chat_id,)).fetchall()
            ids = [row[0] for row in participants]
            if any(not isinstance(p, str) or not p for p in ids) or len(ids) != len(set(ids)) or set(ids) != set(allowed[thread_id]):
                raise BridgeError('Thread participants changed or unavailable')
            sql = '''SELECT m.guid, m.text, m.date, m.is_from_me, h.id
                     FROM chat_message_join j JOIN message m ON m.ROWID=j.message_id
                     LEFT JOIN handle h ON h.ROWID=m.handle_id
                     WHERE j.chat_id=? AND m.text IS NOT NULL'''
            args = [chat_id]
            if query is not None:
                sql += ' AND instr(m.text, ?) > 0'
                args.append(query)
            sql += ' ORDER BY m.date DESC, m.ROWID DESC LIMIT ?'
            args.append(limit)
            result = []
            for guid, body, date, outgoing, sender in con.execute(sql, args):
                if not isinstance(guid, str) or not guid or not isinstance(body, str) or type(date) is not int or date <= 0 or outgoing not in (0, 1) or type(outgoing) is not int:
                    raise BridgeError('Malformed message record')
                if outgoing == 0 and (not isinstance(sender, str) or sender not in ids):
                    raise BridgeError('Message sender unavailable')
                try:
                    timestamp = (EPOCH + timedelta(microseconds=date // 1000)).astimezone(ZoneInfo('America/Chicago')).isoformat()
                except (OverflowError, ValueError):
                    raise BridgeError('Malformed message record') from None
                result.append({'thread_id': thread_id, 'message_id': guid, 'sender': 'self' if outgoing else sender, 'direction': 'sent' if outgoing else 'received', 'timestamp': timestamp, 'text': body})
            return result
    except sqlite3.Error:
        raise BridgeError('Messages database unavailable or unsupported') from None
    finally:
        con.close()
