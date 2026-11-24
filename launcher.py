import os
import sys
import time
import subprocess
import webbrowser

PORT = 8502

def kill_process_on_port(port: int):
    """Find and terminate any process listening on the given TCP port."""
    try:
        # Query netstat for listening processes on this port
        output = subprocess.check_output(
            f'netstat -ano -p tcp | findstr ":{port} "',
            shell=True,
            text=True,
            stderr=subprocess.DEVNULL
        )
        pids = set()
        for line in output.strip().splitlines():
            parts = line.split()
            if len(parts) >= 5 and "LISTENING" in parts[3].upper():
                pids.add(parts[-1])
        
        current_pid = str(os.getpid())
        for pid in pids:
            if pid and pid != "0" and pid != current_pid:
                print(f"[!] Releasing port {port} from existing instance (PID {pid})...")
                subprocess.run(f"taskkill /F /PID {pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(0.5)
    except Exception:
        # Port is already free
        pass

def main():
    print("=" * 60)
    print("  Abstract of Canvass Generator - Municipality of Maragondon")
    print("=" * 60)
    print(f"\n[*] Checking port {PORT} availability...")
    kill_process_on_port(PORT)
    time.sleep(1)

    url = f"http://localhost:{PORT}"
    print(f"[*] Starting application on {url} ...")
    print("[*] Press Ctrl+C in this window anytime to stop the server.\n")

    # Launch Streamlit process
    cmd = [sys.executable, "-m", "streamlit", "run", "app.py", f"--server.port={PORT}"]
    try:
        # Open browser after a brief delay so the server has initialized
        def open_browser():
            time.sleep(2)
            try:
                webbrowser.open(url)
            except Exception:
                pass

        import threading
        t = threading.Thread(target=open_browser, daemon=True)
        t.start()

        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\n[*] Server stopped.")

if __name__ == "__main__":
    main()
