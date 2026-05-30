"""CDM Toolkit CLI entry point."""
import sys
import os


def main():
    """Entry point for cdm-tools command."""
    if len(sys.argv) < 2:
        print("CDM Toolkit v3.0")
        print("Usage: cdm-tools serve [--port PORT] [--no-browser]")
        print("       cdm-tools version")
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "version":
        from cdm_engine import __version__
        print(f"CDM Toolkit v{__version__}")
    elif cmd == "serve":
        port = 8520
        open_browser = True
        i = 2
        while i < len(sys.argv):
            if sys.argv[i] == "--port" and i + 1 < len(sys.argv):
                port = int(sys.argv[i + 1])
                i += 2
            elif sys.argv[i] == "--no-browser":
                open_browser = False
                i += 1
            else:
                i += 1
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        os.chdir(project_root)
        from server import start_server
        start_server(port=port, open_browser=open_browser)
    else:
        print(f"Unknown command: {cmd}")
        sys.exit(1)


if __name__ == "__main__":
    main()
