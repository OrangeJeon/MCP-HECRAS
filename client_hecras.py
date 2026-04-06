from mcp.server.fastmcp import FastMCP
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
import asyncio

async def main():
   server_params = StdioServerParameters(
       command="C:\\python\\.venv\\Scripts\\python.exe",
        args=["C:/python/mcp/server_hecras.py"]
   )
   async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as Session:
            await Session.initialize()

            tools = await Session.list_tools()
            print("Tools: ", [t.name for t in tools.tools])

            result3 = await Session.call_tool("check_connection", {})
            print("check_connection: ", result3)


            result = await Session.call_tool(
                "open_project",
                {"project_path": "C:/Users/kwater/Desktop/RAS/beforeLSB/hwangriver.prj"}
            )
            
            print("Result: ", result)
            
            result4 = await Session.call_tool("get_steady_flow_data", {})
            print("get_result: ", result4)
            
            result2 = await Session.call_tool("run_current_plan", {})
            print("Run Result: ", result2)

            

asyncio.run(main())