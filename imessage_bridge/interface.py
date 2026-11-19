"""MCP tool definitions and strict dispatcher; transport intentionally absent."""
from . import bridge

TOOLS = [
    {'name': 'list_allowed_message_threads', 'description': 'List configured thread identifiers only.', 'inputSchema': {'type': 'object', 'properties': {}, 'additionalProperties': False}, 'annotations': {'readOnlyHint': True, 'openWorldHint': False}},
    {'name': 'read_recent_messages', 'description': 'Read bounded recent plain-text messages in an allowed thread.', 'inputSchema': {'type': 'object', 'properties': {'thread_id': {'type': 'string'}, 'limit': {'type': 'integer', 'minimum': 1, 'maximum': 50}}, 'required': ['thread_id'], 'additionalProperties': False}, 'annotations': {'readOnlyHint': True, 'openWorldHint': False}},
    {'name': 'search_allowed_messages', 'description': 'Search plain-text messages in an allowed thread.', 'inputSchema': {'type': 'object', 'properties': {'thread_id': {'type': 'string'}, 'query': {'type': 'string', 'minLength': 1, 'maxLength': 200}, 'limit': {'type': 'integer', 'minimum': 1, 'maximum': 50}}, 'required': ['thread_id', 'query'], 'additionalProperties': False}, 'annotations': {'readOnlyHint': True, 'openWorldHint': False}},
]


def dispatch(name, arguments):
    definitions = {tool['name']: tool for tool in TOOLS}
    if name not in definitions or not isinstance(arguments, dict):
        raise bridge.BridgeError('Invalid request')
    schema = definitions[name]['inputSchema']
    if set(arguments) - set(schema['properties']) or not set(schema.get('required', [])) <= set(arguments):
        raise bridge.BridgeError('Invalid request')
    return getattr(bridge, name)(**arguments)
