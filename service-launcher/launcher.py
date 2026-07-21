"""
launcher.py
-----------
Service registry manager for all local apps and servers.

Commands:
    python launcher.py status        - Show all services and their port status
    python launcher.py start         - Start all autostart=true services
    python launcher.py start all     - Start every service
    python launcher.py start <name>  - Start a specific service by name/tag
    python launcher.py scan          - Scan ALL listening ports on this machine
    python launcher.py stop <port>   - Kill process on a specific port
"""

import os
import sys
import json
import socket
import subprocess
import psutil
from datetime import datetime

SERVICES_FILE = os.path.join(os.path.dirname(__file__), "services.json")
NGROK_EXE = r"C:\Users\dell\AppData\Local\Microsoft\WinGet\Packages\Ngrok.Ngrok_Microsoft.Winget.Source_8wekyb3d8bbwe\ngrok.exe"

# ─── Colors ──────────────────────────────────────────────────────────────────
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
BOLD   = "\033[1m"
RESET  = "\033[0m"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def load_services():
    with open(SERVICES_FILE, "r") as f:
        return json.load(f)["services"]

def is_port_open(port):
    """Check if a port is currently listening."""
    if port is None:
        return None
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", port)) == 0

def get_pid_on_port(port):
    """Return PID of process listening on a port."""
    for conn in psutil.net_connections(kind="inet"):
        if conn.laddr.port == port and conn.status == "LISTEN":
            return conn.pid
    return None

def is_service_running(svc):
    """Check if a service is running — by port or by process name for port-less services."""
    port = svc.get("port")
    if port is not None:
        return is_port_open(port)
    # For port-less services, check if the script is running as a process
    script = svc["cmd"].split()[-1]  # e.g. screenshot_to_xls.py
    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            cmdline = " ".join(proc.info["cmdline"] or [])
            if script in cmdline:
                return True
        except Exception:
            pass
    return False


    """Return all ports currently in LISTEN state with process info."""
    results = []
    for conn in psutil.net_connections(kind="inet"):
        if conn.status == "LISTEN":
            pid  = conn.pid
            port = conn.laddr.port
            addr = conn.laddr.ip
            try:
                proc = psutil.Process(pid)
                name = proc.name()
                cmd  = " ".join(proc.cmdline())[:80]
            except Exception:
                name = "Unknown"
                cmd  = ""
            results.append({
                "port": port, "addr": addr,
                "pid": pid, "name": name, "cmd": cmd
            })
    return sorted(results, key=lambda x: x["port"])

# ─── Status ──────────────────────────────────────────────────────────────────

def cmd_status():
    services = load_services()
    print(f"\n{BOLD}{'='*70}{RESET}")
    print(f"{BOLD}  SERVICE REGISTRY STATUS  —  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}{RESET}")
    print(f"{BOLD}{'='*70}{RESET}\n")

    fmt = "{:<3} {:<30} {:<8} {:<10} {}"
    print(CYAN + fmt.format("#", "Service", "Port", "Status", "Tags") + RESET)
    print("-" * 70)

    for i, svc in enumerate(services, 1):
        port   = svc.get("port")
        active = is_service_running(svc)
        auto   = "[A]" if svc.get("autostart") else "[ ]"

        if port is None and active:
            status = f"{GREEN}RUNNING{RESET}"
        elif port is None:
            status = f"{RED}STOPPED{RESET}"
        elif active:
            pid    = get_pid_on_port(port)
            status = f"{GREEN}RUNNING (PID {pid}){RESET}"
        else:
            status = f"{RED}STOPPED{RESET}"

        port_str = str(svc.get("port")) if svc.get("port") else "-"
        tags     = ", ".join(svc.get("tags", []))
        print(fmt.format(f"{auto}{i}", svc["name"][:29], port_str, "", tags))
        print(f"   {' '*30}  {status}")

    print()

# ─── Scan all ports ──────────────────────────────────────────────────────────

def cmd_scan():
    print(f"\n{BOLD}{'='*80}{RESET}")
    print(f"{BOLD}  ALL LISTENING PORTS ON THIS MACHINE  —  {datetime.now().strftime('%H:%M:%S')}{RESET}")
    print(f"{BOLD}{'='*80}{RESET}\n")

    ports = get_all_listening_ports()
    fmt   = "{:<8} {:<20} {:<8} {}"
    print(CYAN + fmt.format("PORT", "PROCESS", "PID", "COMMAND") + RESET)
    print("-" * 80)

    for p in ports:
        print(fmt.format(p["port"], p["name"][:19], p["pid"], p["cmd"][:50]))

    print(f"\n{len(ports)} ports listening.\n")

# ─── Start a service ─────────────────────────────────────────────────────────

def start_service(svc):
    name = svc["name"]
    port = svc.get("port")
    cwd  = svc["cwd"]
    cmd  = svc["cmd"]

    # Already running?
    if port and is_port_open(port):
        pid = get_pid_on_port(port)
        print(f"  {YELLOW}⚠  {name} already running on port {port} (PID {pid}){RESET}")
        return

    if not os.path.exists(cwd):
        print(f"  {RED}✗  {name} — working dir not found: {cwd}{RESET}")
        return

    # Build full command
    full_cmd = cmd
    if svc["type"] == "exe" and not os.path.isabs(cmd.split()[0]):
        exe = os.path.join(cwd, cmd.split()[0])
        if os.path.exists(exe):
            full_cmd = exe + " " + " ".join(cmd.split()[1:])

    print(f"  {CYAN}▶  Starting: {name}{RESET}")
    print(f"     CMD : {full_cmd}")
    print(f"     CWD : {cwd}")

    # Launch in new visible terminal window
    subprocess.Popen(
        ["powershell", "-NoExit", "-Command",
         f'cd "{cwd}"; {full_cmd}'],
        creationflags=subprocess.CREATE_NEW_CONSOLE
    )

    port_str = f"port {port}" if port else "no port"
    print(f"  {GREEN}✓  Launched ({port_str}){RESET}\n")

def cmd_start(target="autostart"):
    services = load_services()
    print(f"\n{BOLD}Starting services...{RESET}\n")

    for svc in services:
        if target == "autostart" and not svc.get("autostart"):
            continue
        if target == "all":
            pass
        elif target not in ("autostart", "all"):
            # match by name or tag
            match = (
                target.lower() in svc["name"].lower() or
                target.lower() in [t.lower() for t in svc.get("tags", [])]
            )
            if not match:
                continue
        start_service(svc)

# ─── Stop by port ─────────────────────────────────────────────────────────────

def cmd_stop(port):
    port = int(port)
    pid  = get_pid_on_port(port)
    if pid:
        try:
            psutil.Process(pid).terminate()
            print(f"{GREEN}✓  Killed PID {pid} on port {port}{RESET}")
        except Exception as e:
            print(f"{RED}✗  Could not kill PID {pid}: {e}{RESET}")
    else:
        print(f"{YELLOW}No process found on port {port}{RESET}")

# ─── Main ─────────────────────────────────────────────────────────────────────

def print_help():
    print(f"""
{BOLD}Service Launcher — Usage:{RESET}

  python launcher.py {CYAN}status{RESET}           Show all services and port status
  python launcher.py {CYAN}scan{RESET}             Scan ALL listening ports on this machine
  python launcher.py {CYAN}start{RESET}            Start all autostart services
  python launcher.py {CYAN}start all{RESET}        Start every service
  python launcher.py {CYAN}start <name/tag>{RESET} Start matching service (e.g. start wati)
  python launcher.py {CYAN}stop <port>{RESET}      Kill process on a specific port
""")

if __name__ == "__main__":
    os.system("color")  # enable ANSI colors on Windows
    args = sys.argv[1:]

    if not args or args[0] == "status":
        cmd_status()
    elif args[0] == "scan":
        cmd_scan()
    elif args[0] == "start":
        target = args[1] if len(args) > 1 else "autostart"
        cmd_start(target)
    elif args[0] == "stop" and len(args) > 1:
        cmd_stop(args[1])
    else:
        print_help()
