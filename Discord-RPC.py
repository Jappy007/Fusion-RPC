# Discord Rich Presence Add-In for Autodesk Fusion
# Updated for Fusion (formerly Fusion 360) and latest Discord RPC
# Python 3.12.4 compatible

import adsk.core
import adsk.fusion
import adsk.cam
import traceback
import os
import sys
import time
import threading

# Add current directory to path for module imports
_app = adsk.core.Application.get()
_ui = _app.userInterface

# Get the directory of this script
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

# Import pypresence from local modules folder
try:
    from .modules.pypresence import Presence
except ImportError as e:
    _ui.messageBox(f'Failed to import pypresence: {str(e)}\n\nMake sure the modules/pypresence folder exists.')
    raise

# Import config
try:
    import config
except ImportError as e:
    _ui.messageBox(f'Failed to import config: {str(e)}\n\nMake sure config.py exists.')
    raise

# Global variables
_handlers = []
_rpc = None
_rpc_thread = None
_rpc_running = False
_start_time = None


class DiscordRPCThread(threading.Thread):
    """Background thread to handle Discord RPC updates"""
    
    def __init__(self):
        super().__init__(daemon=True)
        self.client_id = config.CLIENT_ID
        self.update_interval = config.UPDATE_INTERVAL
        self.rpc = None
        
    def run(self):
        """Main thread loop"""
        global _rpc_running, _start_time
        
        try:
            # Initialize Discord RPC
            self.rpc = Presence(self.client_id)
            self.rpc.connect()
            _start_time = int(time.time())
            
            print('Discord Rich Presence connected!')
            
            # Main update loop
            while _rpc_running:
                try:
                    self.update_presence()
                    time.sleep(self.update_interval)
                except Exception as e:
                    print(f'Error updating presence: {str(e)}')
                    time.sleep(self.update_interval)
                    
        except Exception as e:
            print(f'Failed to connect to Discord: {str(e)}')
            print(f'Make sure Discord is running and CLIENT_ID is correct.')
        finally:
            if self.rpc:
                try:
                    self.rpc.close()
                except:
                    pass
    
    def update_presence(self):
        """Update Discord Rich Presence with current Fusion state"""
        try:
            app = adsk.core.Application.get()

            # Get project and document name
            doc = app.activeDocument
            if not doc:
                details = "Idle"
            else:
                # Get project name
                project_name = "Unknown Project"
                try:
                    if doc.dataFile:
                        project = doc.dataFile.parentProject
                        if project:
                            project_name = project.name
                except Exception:
                    pass
                
                # Get document name (remove .f3d extension if present)
                doc_name = doc.name
                if doc_name.endswith('.f3d'):
                    doc_name = doc_name[:-4]
                
                details = f"Project: {project_name}"
                state = f"Working on: {doc_name}"
            
            # Update presence
            self.rpc.update(
                details=details,
                state=state,
                large_image=config.LARGE_IMAGE,
                large_text=config.LARGE_TEXT,
                small_image=config.SMALL_IMAGE if hasattr(config, 'SMALL_IMAGE') else None,
                small_text=config.SMALL_TEXT if hasattr(config, 'SMALL_TEXT') else None,
                start=_start_time
            )
        except: 
            pass


def run(context):
    """Called when the add-in is started"""
    global _rpc_thread, _rpc_running
    
    try:
        # Start Discord RPC thread
        _rpc_running = True
        _rpc_thread = DiscordRPCThread()
        _rpc_thread.start()
        
    except Exception as e:
        _ui.messageBox(f'Failed to start:\n{traceback.format_exc()}')


def stop(context):
    """Called when the add-in is stopped"""
    global _rpc_running, _rpc_thread, _rpc
    
    try:
        # Stop RPC thread
        _rpc_running = False
        
        # Clear Discord presence immediately
        if _rpc_thread and hasattr(_rpc_thread, 'rpc') and _rpc_thread.rpc:
            try:
                _rpc_thread.rpc.clear()
            except:
                pass
        
        # Give thread time to clean up
        if _rpc_thread and _rpc_thread.is_alive():
            _rpc_thread.join(timeout=2.0)
        
        print('Discord RPC stopped and cleared.')
        
    except Exception as e:
        print(f'Failed to stop:\n{traceback.format_exc()}')