import typer
from pathlib import Path

from .server import run_sse, run_stdio, run_streamable_http

def validate_excel_path(path: str | None) -> Path | None:
    """Validate Excel directory path"""
    if path is None:
        return None
    excel_path = Path(path)
    if not excel_path.exists():
        raise typer.BadParameter(f"Path '{path}' does not exist.")
    return excel_path

app = typer.Typer(help="Excel MCP Server")

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    try:
        run_sse()
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    try:
        run_streamable_http()
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()


@app.command()
def stdio(excel_files_path: str | None = typer.Option(
        None, help="""
        Path to Excel files directory. Must exist if defined.
        If defined, the mcp server will be limited to the files in this directory. And works with relative paths.
        If not defined, the mcp server will work with absolute paths.
        """
    )):
    """Start Excel MCP Server in stdio mode"""
    excel_path = validate_excel_path(excel_files_path)
    try:
        run_stdio(excel_files_path=excel_path)
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    app() 