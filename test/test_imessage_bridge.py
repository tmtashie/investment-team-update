import contextlib
import io
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from imessage_bridge import bridge, interface


class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.db = Path(self.tmp.name) / 'chat.db'
        self.config = Path(self.tmp.name) / 'allowlist.json'
        self.config.write_text('{"threads":{"allowed":["a", "b"]}}')
        self.p1 = patch.object(bridge, 'DB', self.db)
        self.p2 = patch.object(bridge, 'CONFIG', self.config)
        self.p1.start(); self.p2.start()
        self.addCleanup(self.p1.stop); self.addCleanup(self.p2.stop)
        con = sqlite3.connect(self.db)
        con.executescript('''CREATE TABLE chat (guid TEXT); CREATE TABLE handle (id TEXT);
          CREATE TABLE chat_handle_join (chat_id INTEGER, handle_id INTEGER);
          CREATE TABLE message (guid TEXT, text TEXT, date INTEGER, is_from_me INTEGER, handle_id INTEGER);
          CREATE TABLE chat_message_join (chat_id INTEGER, message_id INTEGER);
          INSERT INTO chat VALUES ('allowed'), ('private');
          INSERT INTO handle VALUES ('a'), ('b'), ('c');
          INSERT INTO chat_handle_join VALUES (1,1),(1,2),(2,3);
          INSERT INTO message VALUES ('m1','first',1000000000,0,1), ('m2','second',2000000000,1,0), ('m3','secret',3000000000,0,3);
          INSERT INTO chat_message_join VALUES (1,1),(1,2),(2,3);''')
        con.commit(); con.close()

    def test_authorization_order_direction_and_search(self):
        self.assertEqual(bridge.list_allowed_message_threads(), [{'thread_id': 'allowed'}])
        rows = bridge.read_recent_messages('allowed')
        self.assertEqual([r['text'] for r in rows], ['second', 'first'])
        self.assertEqual([r['direction'] for r in rows], ['sent', 'received'])
        self.assertEqual(rows[1]['sender'], 'a')
        self.assertTrue(rows[0]['timestamp'].endswith('-06:00'))
        self.assertEqual(len(bridge.search_allowed_messages('allowed', 'first')), 1)
        with self.assertRaises(bridge.BridgeError): bridge.read_recent_messages('private')

    def test_group_membership_change_fails_closed(self):
        con = sqlite3.connect(self.db); con.execute('INSERT INTO chat_handle_join VALUES (1,3)'); con.commit(); con.close()
        with self.assertRaises(bridge.BridgeError): bridge.read_recent_messages('allowed')

    def test_membership_and_messages_share_one_snapshot(self):
        con = sqlite3.connect(self.db)
        con.execute('PRAGMA journal_mode=WAL')
        con.close()
        connect = bridge._connect
        changed = []

        def connect_with_concurrent_change():
            reader = connect()

            def mutate_before_message_query(statement):
                if 'SELECT m.guid' not in statement or changed:
                    return
                changed.append(True)
                writer = sqlite3.connect(self.db)
                writer.execute('INSERT INTO chat_handle_join VALUES (1,3)')
                writer.execute("INSERT INTO message VALUES ('m4','new private content',4000000000,1,0)")
                writer.execute('INSERT INTO chat_message_join VALUES (1,4)')
                writer.commit()
                writer.close()

            reader.set_trace_callback(mutate_before_message_query)
            return reader

        with patch.object(bridge, '_connect', side_effect=connect_with_concurrent_change):
            rows = bridge.read_recent_messages('allowed')
        self.assertTrue(changed)
        self.assertEqual([row['text'] for row in rows], ['second', 'first'])
        with self.assertRaises(bridge.BridgeError):
            bridge.read_recent_messages('allowed')

    def test_no_sql_paths_or_write_operations(self):
        for args in ({'thread_id': 'allowed', 'sql': 'DELETE FROM message'}, {'thread_id': 'allowed', 'path': '/tmp/x'}):
            with self.assertRaises(bridge.BridgeError): interface.dispatch('read_recent_messages', args)
        with self.assertRaises(bridge.BridgeError): interface.dispatch('send_message', {})
        self.assertTrue(all(t['annotations']['readOnlyHint'] for t in interface.TOOLS))
        con = bridge._connect()
        with self.assertRaises(sqlite3.OperationalError): con.execute("DELETE FROM message")
        con.close()

    def test_malformed_and_no_body_logging(self):
        con = sqlite3.connect(self.db); con.execute("UPDATE message SET date='bad' WHERE guid='m1'"); con.commit(); con.close()
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr):
            with self.assertRaises(bridge.BridgeError) as error: bridge.read_recent_messages('allowed')
        self.assertNotIn('first', str(error.exception) + stderr.getvalue())
        self.assertNotIn('second', str(error.exception) + stderr.getvalue())


if __name__ == '__main__': unittest.main()

