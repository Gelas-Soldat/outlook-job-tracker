import asyncio
import requests
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

BASE_URL = 'http://127.0.0.1:8080'

server = Server('outlook-mail')

@server.list_tools()
async def list_tools():
    return [
        types.Tool(
            name='search_emails',
            description='Search Ryan emails by keyword',
            inputSchema={
                'type': 'object',
                'properties': {
                    'query': {'type': 'string', 'description': 'Search keyword'},
                    'top': {'type': 'integer', 'description': 'Number of results', 'default': 25}
                },
                'required': ['query']
            }
        ),
        types.Tool(
            name='get_job_applications',
            description='Get all job application emails since January 2026',
            inputSchema={
                'type': 'object',
                'properties': {
                    'after': {'type': 'string', 'description': 'Start date e.g. 2026-01-01T00:00:00Z'}
                }
            }
        ),
        types.Tool(
            name='get_recent_emails',
            description='Get most recent emails from inbox',
            inputSchema={
                'type': 'object',
                'properties': {
                    'top': {'type': 'integer', 'description': 'Number of emails to return', 'default': 20}
                }
            }
        ),
        types.Tool(
            name='read_email',
            description='Read the full content of a specific email by ID',
            inputSchema={
                'type': 'object',
                'properties': {
                    'email_id': {'type': 'string', 'description': 'The email message ID'}
                },
                'required': ['email_id']
            }
        ),
        types.Tool(
            name='get_calendar_events',
            description='Get upcoming calendar events',
            inputSchema={
                'type': 'object',
                'properties': {}
            }
        )
    ]

@server.call_tool()
async def call_tool(name: str, arguments: dict):
    try:
        if name == 'search_emails':
            q = arguments.get('query', '')
            top = arguments.get('top', 25)
            resp = requests.get(f'{BASE_URL}/emails/search', params={'q': q, 'top': top})
            return [types.TextContent(type='text', text=resp.text)]

        elif name == 'get_job_applications':
            after = arguments.get('after', '2026-01-01T00:00:00Z')
            resp = requests.get(f'{BASE_URL}/emails/jobs', params={'after': after, 'top': 200})
            return [types.TextContent(type='text', text=resp.text)]

        elif name == 'get_recent_emails':
            top = arguments.get('top', 20)
            resp = requests.get(f'{BASE_URL}/emails/recent', params={'top': top})
            return [types.TextContent(type='text', text=resp.text)]

        elif name == 'read_email':
            email_id = arguments.get('email_id', '')
            resp = requests.get(f'{BASE_URL}/emails/read/{email_id}')
            return [types.TextContent(type='text', text=resp.text)]

        elif name == 'get_calendar_events':
            resp = requests.get(f'{BASE_URL}/calendar/events')
            return [types.TextContent(type='text', text=resp.text)]

        else:
            return [types.TextContent(type='text', text=f'Unknown tool: {name}')]

    except Exception as e:
        return [types.TextContent(type='text', text=f'Error: {str(e)}')]

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())

if __name__ == '__main__':
    asyncio.run(main())
