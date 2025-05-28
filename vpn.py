#!/usr/bin/env python3
"""
Windows WireGuard Tunnel Manager
Handles WireGuard configuration on Windows systems
"""

import subprocess
import sys
import os
import tempfile
import platform
import time
import socket
import requests
from pathlib import Path
import re
import winreg
import ctypes

class WindowsWireGuardTunnel:
    def __init__(self, config_data):
        self.full_config = config_data
        self.interface_name = "wg0"
        self.config_file = None
        self.wireguard_path = None
        self.tunnel_service_name = f"WireGuardTunnel${self.interface_name}"
        
        # Parse configuration
        self.parse_config()
        
    def is_admin(self):
        """Check if running with administrator privileges"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def find_wireguard_installation(self):
        """Find WireGuard installation paths on Windows"""
        possible_paths = [
            r"C:\Program Files\WireGuard\wg.exe",
            r"C:\Program Files (x86)\WireGuard\wg.exe",
            r"C:\Tools\WireGuard\wg.exe",
            r"C:\WireGuard\wg.exe"
        ]
        
        # Check common installation paths
        for path in possible_paths:
            if os.path.exists(path):
                self.wireguard_path = os.path.dirname(path)
                print(f"✓ Found WireGuard at: {self.wireguard_path}")
                return True
        
        # Check Windows PATH
        try:
            result = subprocess.run(["wg", "--version"], capture_output=True, text=True)
            if result.returncode == 0:
                # WireGuard is in PATH
                self.wireguard_path = ""  # Use system PATH
                print("✓ WireGuard found in system PATH")
                return True
        except FileNotFoundError:
            pass
        
        # Check Windows Registry
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WireGuard") as key:
                install_path = winreg.QueryValueEx(key, "InstallPath")[0]
                if os.path.exists(os.path.join(install_path, "wg.exe")):
                    self.wireguard_path = install_path
                    print(f"✓ Found WireGuard via registry: {self.wireguard_path}")
                    return True
        except (FileNotFoundError, OSError):
            pass
        
        print("✗ WireGuard not found")
        self.print_windows_installation_instructions()
        return False
    
    def print_windows_installation_instructions(self):
        """Print WireGuard installation instructions for Windows"""
        print("\n" + "="*60)
        print("WireGuard Installation Required")
        print("="*60)
        print("Please install WireGuard for Windows:")
        print("1. Download from: https://www.wireguard.com/install/")
        print("2. Run the installer as Administrator")
        print("3. Restart this script after installation")
        print("\nAlternative installation methods:")
        print("• Chocolatey: choco install wireguard")
        print("• Scoop: scoop install wireguard")
        print("• Winget: winget install WireGuard.WireGuard")
    
    def get_wg_command(self, cmd):
        """Get full path to WireGuard command"""
        if self.wireguard_path:
            return os.path.join(self.wireguard_path, cmd)
        return cmd
    
    def parse_config(self):
        """Parse WireGuard configuration and extract components"""
        self.interface_config = {}
        self.peer_config = {}
        
        current_section = None
        lines = self.full_config.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
                
            if line.startswith('[') and line.endswith(']'):
                current_section = line[1:-1].lower()
                continue
            
            if '=' in line:
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip()
                
                if current_section == 'interface':
                    self.interface_config[key] = value
                elif current_section == 'peer':
                    self.peer_config[key] = value
        
        print(f"✓ Parsed config - Interface keys: {list(self.interface_config.keys())}")
        print(f"✓ Parsed config - Peer keys: {list(self.peer_config.keys())}")
    
    def get_real_ip(self):
        """Get the current real IP address"""
        try:
            response = requests.get("https://api.ipify.org", timeout=10)
            return response.text.strip()
        except:
            try:
                response = requests.get("https://ifconfig.me", timeout=10)
                return response.text.strip()
            except:
                return "Unable to determine"
    
    def create_config_file(self):
        """Create WireGuard configuration file for Windows"""
        try:
            # Create config in Windows temp directory
            temp_dir = os.environ.get('TEMP', r'C:\temp')
            self.config_file = os.path.join(temp_dir, f"{self.interface_name}.conf")
            
            with open(self.config_file, 'w') as f:
                f.write(self.full_config.strip())
            
            print(f"✓ Configuration file created: {self.config_file}")
            return True
        except Exception as e:
            print(f"✗ Failed to create config file: {e}")
            return False
    
    def get_windows_interface_name(self):
        """Get the Windows network interface name for WireGuard"""
        try:
            # On Windows, WireGuard creates interfaces with specific naming
            # The interface name in Windows is usually the tunnel name
            return self.interface_name
        except Exception as e:
            print(f"Error getting interface name: {e}")
            return self.interface_name
    
    def start_tunnel_windows_service(self):
        """Start WireGuard tunnel using Windows service method"""
        try:
            print("🔄 Starting WireGuard tunnel (Windows service method)...")
            
            # Method 1: Use wireguard.exe to install and start service
            wg_exe = self.get_wg_command("wireguard.exe")
            
            if os.path.exists(wg_exe):
                # Install tunnel configuration
                result = subprocess.run([
                    wg_exe, "/installtunnelservice", self.config_file
                ], capture_output=True, text=True)
                
                if result.returncode == 0:
                    print("✓ Tunnel service installed successfully")
                    return True
                else:
                    print(f"⚠ Service install warning: {result.stderr}")
                    return self.start_tunnel_wg_quick()
            else:
                return self.start_tunnel_wg_quick()
                
        except Exception as e:
            print(f"✗ Service method failed: {e}")
            return self.start_tunnel_wg_quick()
    
    def start_tunnel_wg_quick(self):
        """Start WireGuard tunnel using wg-quick method"""
        try:
            print("🔄 Starting WireGuard tunnel (wg-quick method)...")
            
            wg_quick = self.get_wg_command("wg-quick.exe")
            
            # Try wg-quick up
            result = subprocess.run([
                wg_quick, "up", self.config_file
            ], capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✓ WireGuard tunnel started successfully")
                return True
            else:
                print(f"✗ wg-quick failed: {result.stderr}")
                return self.start_tunnel_manual_windows()
                
        except Exception as e:
            print(f"✗ wg-quick method failed: {e}")
            return self.start_tunnel_manual_windows()
    
    def start_tunnel_manual_windows(self):
        """Manually start WireGuard tunnel on Windows"""
        try:
            print("🔧 Setting up tunnel manually on Windows...")
            
            # This is more complex on Windows and requires administrative tools
            # For now, we'll use PowerShell commands
            
            # Create the interface using netsh or PowerShell
            powershell_script = f"""
# Create WireGuard interface
$interfaceName = "{self.interface_name}"
$configFile = "{self.config_file}"

# Add the tunnel using WireGuard
& '{self.get_wg_command("wg.exe")}' setconf $interfaceName $configFile
"""
            
            # Execute PowerShell script
            result = subprocess.run([
                "powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_script
            ], capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✓ Manual tunnel setup completed")
                return True
            else:
                print(f"✗ Manual setup failed: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"✗ Manual Windows setup failed: {e}")
            return False
    
    def check_tunnel_status_windows(self):
        """Check WireGuard tunnel status on Windows"""
        try:
            wg_cmd = self.get_wg_command("wg.exe")
            result = subprocess.run([wg_cmd, "show"], capture_output=True, text=True)
            
            if result.returncode == 0 and self.interface_name in result.stdout:
                return True
            
            # Also check Windows services
            result = subprocess.run([
                "sc", "query", self.tunnel_service_name
            ], capture_output=True, text=True)
            
            return "RUNNING" in result.stdout
            
        except Exception:
            return False
    
    def stop_tunnel_windows(self):
        """Stop WireGuard tunnel on Windows"""
        try:
            print("🛑 Stopping WireGuard tunnel...")
            
            # Method 1: Stop service
            subprocess.run([
                "sc", "stop", self.tunnel_service_name
            ], capture_output=True)
            
            # Method 2: Use wg-quick
            wg_quick = self.get_wg_command("wg-quick.exe")
            subprocess.run([
                wg_quick, "down", self.config_file
            ], capture_output=True)
            
            # Method 3: Use wireguard.exe
            wg_exe = self.get_wg_command("wireguard.exe")
            if os.path.exists(wg_exe):
                subprocess.run([
                    wg_exe, "/uninstalltunnelservice", self.interface_name
                ], capture_output=True)
            
            print("✓ Tunnel stopped")
            return True
            
        except Exception as e:
            print(f"⚠ Error stopping tunnel: {e}")
            return False
    
    def test_connection(self):
        """Test the VPN connection"""
        print("\n🔍 Testing VPN connection...")
        
        # Wait for connection to establish
        time.sleep(5)
        
        # Test IP change
        new_ip = self.get_real_ip()
        print(f"New IP: {new_ip}")
        
        # Check if IP changed from original
        if new_ip == "34.102.88.164":
            print("✅ SUCCESS! IP changed to VPN server - VPN is working!")
        elif new_ip != "49.145.198.195":
            print(f"✅ IP changed to: {new_ip} - VPN may be working")
        else:
            print("⚠ IP hasn't changed - VPN may not be routing traffic")
            self.diagnose_windows_connection()
        
        # Test DNS resolution
        try:
            socket.gethostbyname("google.com")
            print("✓ DNS resolution working")
        except:
            print("✗ DNS resolution failed")
        
        # Test internet connectivity
        try:
            response = requests.get("https://www.google.com", timeout=10)
            if response.status_code == 200:
                print("✓ Internet connectivity working")
            else:
                print("✗ Internet connectivity issues")
        except Exception as e:
            print(f"✗ Internet connectivity failed: {e}")
    
    def diagnose_windows_connection(self):
        """Diagnose connection issues on Windows"""
        print("\n🔍 Diagnosing Windows connection...")
        
        try:
            # Check WireGuard status
            wg_cmd = self.get_wg_command("wg.exe")
            result = subprocess.run([wg_cmd, "show"], capture_output=True, text=True)
            print("WireGuard interfaces:")
            if result.stdout.strip():
                print(result.stdout)
            else:
                print("  No active interfaces found")
            
            # Check Windows network interfaces
            result = subprocess.run([
                "ipconfig", "/all"
            ], capture_output=True, text=True)
            
            # Look for WireGuard adapter
            if "WireGuard" in result.stdout or self.interface_name in result.stdout:
                print("✓ WireGuard network adapter found")
            else:
                print("✗ WireGuard network adapter not found")
            
            # Check routing table
            result = subprocess.run([
                "route", "print"
            ], capture_output=True, text=True)
            
            print("\nWindows routing table (first 20 lines):")
            lines = result.stdout.split('\n')[:20]
            for line in lines:
                if line.strip():
                    print(f"  {line}")
            
        except Exception as e:
            print(f"Error during diagnosis: {e}")
    
    def cleanup_windows(self):
        """Clean up temporary files on Windows"""
        if self.config_file and os.path.exists(self.config_file):
            try:
                os.unlink(self.config_file)
                print("✓ Temporary config file cleaned up")
            except:
                print(f"⚠ Could not remove temporary file: {self.config_file}")
    
    def run(self):
        """Main execution method for Windows"""
        print("🚀 Windows WireGuard Tunnel Manager")
        print("=" * 50)
        
        # Check administrator privileges
        if not self.is_admin():
            print("✗ This script must be run as Administrator on Windows")
            print("Right-click Command Prompt or PowerShell and 'Run as Administrator'")
            return False
        
        # Find WireGuard installation
        if not self.find_wireguard_installation():
            return False
        
        # Show current IP
        original_ip = self.get_real_ip()
        print(f"\n📍 Original IP: {original_ip}")
        
        # Create config file
        if not self.create_config_file():
            return False
        
        try:
            # Start tunnel
            success = (self.start_tunnel_windows_service() or 
                      self.start_tunnel_wg_quick() or 
                      self.start_tunnel_manual_windows())
            
            if not success:
                print("✗ Failed to start tunnel with all methods")
                return False
            
            # Test connection
            self.test_connection()
            
            # Keep tunnel running
            print(f"\n🔒 VPN tunnel is active!")
            print("Press Ctrl+C to stop the tunnel...")
            
            try:
                while True:
                    time.sleep(30)
                    if not self.check_tunnel_status_windows():
                        print("⚠ Tunnel appears to be down!")
                        # Try to restart
                        print("🔄 Attempting to restart tunnel...")
                        if not self.start_tunnel_wg_quick():
                            break
            except KeyboardInterrupt:
                print("\n\n🛑 Stopping tunnel...")
                self.stop_tunnel_windows()
                
        except Exception as e:
            print(f"✗ Unexpected error: {e}")
            
        finally:
            self.cleanup_windows()
        
        return True

def main():
    # Check if running on Windows
    if platform.system() != "Windows":
        print("This script is designed for Windows. Use the Linux version for other systems.")
        return
    
    # Your WireGuard configuration
    config = """[Interface]
PrivateKey = wLy94QgcI4YdpjXv4qFbJBqTFBlCNJQ4Kkaciu2RNGA=
Address = 10.49.0.4/32
DNS = 172.21.156.241
MTU = 1380

[Peer]
PublicKey = p0WqIapEnqA27uVZT5LKXQxHjrfiFL1kEJpXJcpV3DI=
PresharedKey = PijhVjJgo2lVa7B2LTV7LZm102K0ktKDzjnK3Qfbegk=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 34.102.88.164:443
"""

    tunnel = WindowsWireGuardTunnel(config)
    
    try:
        tunnel.run()
    except KeyboardInterrupt:
        print("\n👋 Exiting...")
        tunnel.cleanup_windows()

if __name__ == "__main__":
    main()