#!/usr/bin/env python
"""
GUI Application for ICP/MP Expert Automation

This application provides a graphical user interface for controlling the 
Agilent ICP/MP Expert instrument using Windows Forms and the Automation SDK.

Features:
- Real-time instrument status display
- Point-and-click control of all instrument functions
- Sample queue management
- Progress monitoring
- Result viewing and export

Requirements:
- IronPython 2.7
- Agilent MP Expert software
"""

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")

import System
from System import *
from System.Drawing import *
from System.Windows.Forms import *
from System.Threading import *
from System.Windows.Forms import Timer
import time
import Automation
from Automation import Automation

class InstrumentControlGUI(Form):
    """Main GUI application for instrument control"""
    
    def __init__(self):
        """Initialize the GUI application"""
        self.client = None
        self.connected = False
        self.status_timer = None
        self.output_path = ""
        self.template_path = ""
        self.lims_output_path = ""  # Separate path for LIMS exports
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main form properties
        self.Text = "ICP/MP Expert Automation Control"
        self.Size = Size(1100, 800)  # Optimized for typical screen sizes (1366x768, 1920x1080)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        
        # Create main layout
        self.create_connection_panel()
        self.create_status_panel()
        self.create_control_panel()
        self.create_worksheet_panel()  # New comprehensive worksheet panel
        self.create_sample_panel()
        self.create_advanced_panel()   # New advanced controls panel
        self.create_log_panel()
        
        # Setup timer for status updates
        self.status_timer = Timer()
        self.status_timer.Interval = 1000  # Update every second
        self.status_timer.Tick += self.update_status
        self.status_timer.Enabled = False  # Start disabled
        
        # Handle form closing
        self.FormClosing += self.form_closing
        
        # Attempt automatic connection at startup
        self.auto_connect_at_startup()
        
    def form_closing(self, sender, e):
        """Handle form closing event"""
        try:
            # Stop the status timer first
            if self.status_timer:
                self.status_timer.Stop()
                self.status_timer.Dispose()
            
            # Disconnect from instrument if connected
            if self.client and self.connected:
                self.log_message("Disconnecting from instrument...")
                
                # Attempt graceful disconnect, ignore socket errors
                try:
                    self.client.Disconnect()
                except Exception as disconnect_ex:
                    # Socket exceptions are expected when connection is already broken
                    if "SocketException" in str(disconnect_ex):
                        self.log_message("Connection already closed by remote host")
                    else:
                        self.log_message("Disconnect warning: {0}".format(str(disconnect_ex)))
                
                # Clean up resources
                try:
                    self.client.Dispose()
                except Exception as dispose_ex:
                    self.log_message("Resource cleanup warning: {0}".format(str(dispose_ex)))
                
                self.connected = False
                self.log_message("Disconnected successfully")
                
        except Exception as ex:
            # Log but don't prevent closing
            print("Error during cleanup: {0}".format(str(ex)))
    
    def auto_connect_at_startup(self):
        """Attempt automatic connection at startup"""
        try:
            # Use default connection values
            host = self.host_textbox.Text
            port_text = self.port_textbox.Text
            
            self.log_message("Attempting automatic connection to {0}:{1}...".format(host, port_text))
            
            # Validate port number
            try:
                port = int(port_text)
                if port < 1 or port > 65535:
                    self.log_message("Invalid default port number, skipping auto-connect")
                    return
            except ValueError:
                self.log_message("Invalid default port format, skipping auto-connect")
                return
            
            # Attempt connection
            self.client = Automation()
            self.client.Connect(host, port)
            
            # Check connection state
            if hasattr(self.client.Client, 'State') and hasattr(self.client.Client.State, 'Connected'):
                connected = (self.client.Client.State.value__ == 1)  # Connected state
            else:
                connected = True  # Assume connected if no exception
            
            if connected:
                self.connected = True
                self.connection_status.Text = "Connected (Auto)"
                self.connection_status.ForeColor = Color.Green
                self.connect_button.Enabled = False
                self.disconnect_button.Enabled = True
                self.enable_controls(True)
                self.status_timer.Start()
                self.log_message("Auto-connection successful")
            else:
                self.log_message("Auto-connection failed - instrument not responding")
                
        except Exception as ex:
            # Auto-connect failed, but don't show error dialogs at startup
            self.log_message("Auto-connection failed: {0}".format(str(ex)))
            self.log_message("You can manually connect using the Connect button")
        
    def create_connection_panel(self):
        """Create connection control panel"""
        connection_group = GroupBox()
        connection_group.Text = "Connection"
        connection_group.Location = Point(10, 10)
        connection_group.Size = Size(300, 80)
        
        # Host/Port inputs
        host_label = Label()
        host_label.Text = "Host:"
        host_label.Location = Point(10, 25)
        host_label.Size = Size(40, 20)
        connection_group.Controls.Add(host_label)
        
        self.host_textbox = TextBox()
        self.host_textbox.Text = "127.0.0.1"
        self.host_textbox.Location = Point(50, 23)
        self.host_textbox.Size = Size(80, 20)
        connection_group.Controls.Add(self.host_textbox)
        
        port_label = Label()
        port_label.Text = "Port:"
        port_label.Location = Point(140, 25)
        port_label.Size = Size(30, 20)
        connection_group.Controls.Add(port_label)
        
        self.port_textbox = TextBox()
        self.port_textbox.Text = "8000"
        self.port_textbox.Location = Point(175, 23)
        self.port_textbox.Size = Size(50, 20)
        connection_group.Controls.Add(self.port_textbox)
        
        # Connect/Disconnect buttons
        self.connect_button = Button()
        self.connect_button.Text = "Connect"
        self.connect_button.Location = Point(10, 50)
        self.connect_button.Size = Size(70, 25)
        self.connect_button.Click += self.connect_clicked
        connection_group.Controls.Add(self.connect_button)
        
        self.disconnect_button = Button()
        self.disconnect_button.Text = "Disconnect"
        self.disconnect_button.Location = Point(90, 50)
        self.disconnect_button.Size = Size(80, 25)
        self.disconnect_button.Enabled = False
        self.disconnect_button.Click += self.disconnect_clicked
        connection_group.Controls.Add(self.disconnect_button)
        
        # Connection status
        self.connection_status = Label()
        self.connection_status.Text = "Not Connected"
        self.connection_status.Location = Point(180, 55)
        self.connection_status.Size = Size(100, 20)
        self.connection_status.ForeColor = Color.Red
        connection_group.Controls.Add(self.connection_status)
        
        self.Controls.Add(connection_group)
    
    def create_status_panel(self):
        """Create instrument status display panel"""
        status_group = GroupBox()
        status_group.Text = "Instrument Status"
        status_group.Location = Point(320, 10)
        status_group.Size = Size(320, 180)  # Reduced width and height for better fit
        
        self.status_listbox = ListBox()
        self.status_listbox.Location = Point(10, 20)
        self.status_listbox.Size = Size(300, 150)  # Adjusted for new panel size
        status_group.Controls.Add(self.status_listbox)
        
        self.Controls.Add(status_group)
    
    def create_control_panel(self):
        """Create instrument control panel"""
        control_group = GroupBox()
        control_group.Text = "Instrument Control"
        control_group.Location = Point(10, 100)
        control_group.Size = Size(300, 250)
        
        # Plasma controls
        plasma_label = Label()
        plasma_label.Text = "Plasma Control:"
        plasma_label.Location = Point(10, 25)
        plasma_label.Size = Size(100, 20)
        control_group.Controls.Add(plasma_label)
        
        self.plasma_on_button = Button()
        self.plasma_on_button.Text = "Ignite"
        self.plasma_on_button.Location = Point(10, 45)
        self.plasma_on_button.Size = Size(60, 30)
        self.plasma_on_button.BackColor = Color.LightGreen
        self.plasma_on_button.Click += self.plasma_on_clicked
        control_group.Controls.Add(self.plasma_on_button)
        
        self.plasma_off_button = Button()
        self.plasma_off_button.Text = "Extinguish"
        self.plasma_off_button.Location = Point(80, 45)
        self.plasma_off_button.Size = Size(80, 30)
        self.plasma_off_button.BackColor = Color.LightCoral
        self.plasma_off_button.Click += self.plasma_off_clicked
        control_group.Controls.Add(self.plasma_off_button)
        
        # Pump controls
        pump_label = Label()
        pump_label.Text = "Pump Control:"
        pump_label.Location = Point(10, 85)
        pump_label.Size = Size(100, 20)
        control_group.Controls.Add(pump_label)
        
        self.pump_off_button = Button()
        self.pump_off_button.Text = "Off"
        self.pump_off_button.Location = Point(10, 105)
        self.pump_off_button.Size = Size(50, 25)
        self.pump_off_button.Click += self.pump_off_clicked
        control_group.Controls.Add(self.pump_off_button)
        
        self.pump_slow_button = Button()
        self.pump_slow_button.Text = "Slow"
        self.pump_slow_button.Location = Point(70, 105)
        self.pump_slow_button.Size = Size(50, 25)
        self.pump_slow_button.Click += self.pump_slow_clicked
        control_group.Controls.Add(self.pump_slow_button)
        
        self.pump_fast_button = Button()
        self.pump_fast_button.Text = "Fast"
        self.pump_fast_button.Location = Point(130, 105)
        self.pump_fast_button.Size = Size(50, 25)
        self.pump_fast_button.Click += self.pump_fast_clicked
        control_group.Controls.Add(self.pump_fast_button)
        
        # Purge controls
        purge_label = Label()
        purge_label.Text = "N2 Purge:"
        purge_label.Location = Point(10, 140)
        purge_label.Size = Size(80, 20)
        control_group.Controls.Add(purge_label)
        
        self.purge_on_button = Button()
        self.purge_on_button.Text = "On"
        self.purge_on_button.Location = Point(10, 160)
        self.purge_on_button.Size = Size(50, 25)
        self.purge_on_button.Click += self.purge_on_clicked
        control_group.Controls.Add(self.purge_on_button)
        
        self.purge_off_button = Button()
        self.purge_off_button.Text = "Off"
        self.purge_off_button.Location = Point(70, 160)
        self.purge_off_button.Size = Size(50, 25)
        self.purge_off_button.Click += self.purge_off_clicked
        control_group.Controls.Add(self.purge_off_button)
        
        # Measurement controls
        measurement_label = Label()
        measurement_label.Text = "Measurement:"
        measurement_label.Location = Point(10, 195)
        measurement_label.Size = Size(100, 20)
        control_group.Controls.Add(measurement_label)
        
        self.start_button = Button()
        self.start_button.Text = "Start"
        self.start_button.Location = Point(10, 215)
        self.start_button.Size = Size(60, 25)
        self.start_button.BackColor = Color.LightBlue
        self.start_button.Click += self.start_clicked
        control_group.Controls.Add(self.start_button)
        
        self.stop_button = Button()
        self.stop_button.Text = "Stop"
        self.stop_button.Location = Point(80, 215)
        self.stop_button.Size = Size(60, 25)
        self.stop_button.BackColor = Color.Orange
        self.stop_button.Click += self.stop_clicked
        control_group.Controls.Add(self.stop_button)
        
        # MP Expert UI controls
        ui_label = Label()
        ui_label.Text = "MP Expert UI:"
        ui_label.Location = Point(170, 195)
        ui_label.Size = Size(100, 20)
        control_group.Controls.Add(ui_label)
        
        self.show_ui_button = Button()
        self.show_ui_button.Text = "Show UI"
        self.show_ui_button.Location = Point(170, 215)
        self.show_ui_button.Size = Size(60, 25)
        self.show_ui_button.BackColor = Color.LightGray
        self.show_ui_button.Click += self.show_ui_clicked
        control_group.Controls.Add(self.show_ui_button)
        
        self.hide_ui_button = Button()
        self.hide_ui_button.Text = "Hide UI"
        self.hide_ui_button.Location = Point(240, 215)
        self.hide_ui_button.Size = Size(60, 25)
        self.hide_ui_button.BackColor = Color.LightGray
        self.hide_ui_button.Click += self.hide_ui_clicked
        control_group.Controls.Add(self.hide_ui_button)

        self.Controls.Add(control_group)
    
    def create_worksheet_panel(self):
        """Create comprehensive worksheet management panel"""
        worksheet_group = GroupBox()
        worksheet_group.Text = "Worksheet Management"
        worksheet_group.Location = Point(320, 200)  # Moved up slightly
        worksheet_group.Size = Size(320, 130)  # Reduced width to match status panel
        
        # Worksheet operations row 1
        self.worksheet_new_button = Button()
        self.worksheet_new_button.Text = "New from Template"
        self.worksheet_new_button.Location = Point(10, 25)
        self.worksheet_new_button.Size = Size(95, 25)  # Slightly smaller to fit
        self.worksheet_new_button.BackColor = Color.LightGreen
        self.worksheet_new_button.Click += self.worksheet_new_clicked
        worksheet_group.Controls.Add(self.worksheet_new_button)
        
        self.worksheet_open_button = Button()
        self.worksheet_open_button.Text = "Open Worksheet"
        self.worksheet_open_button.Location = Point(115, 25)
        self.worksheet_open_button.Size = Size(95, 25)  # Adjusted position and size
        self.worksheet_open_button.BackColor = Color.LightBlue
        self.worksheet_open_button.Click += self.worksheet_open_clicked
        worksheet_group.Controls.Add(self.worksheet_open_button)
        
        self.worksheet_save_button = Button()
        self.worksheet_save_button.Text = "Save As"
        self.worksheet_save_button.Location = Point(220, 25)
        self.worksheet_save_button.Size = Size(75, 25)  # Adjusted size
        self.worksheet_save_button.BackColor = Color.LightYellow
        self.worksheet_save_button.Click += self.worksheet_save_clicked
        worksheet_group.Controls.Add(self.worksheet_save_button)
        
        # Worksheet operations row 2
        self.worksheet_save_close_button = Button()
        self.worksheet_save_close_button.Text = "Save & Close"
        self.worksheet_save_close_button.Location = Point(10, 55)
        self.worksheet_save_close_button.Size = Size(95, 25)
        self.worksheet_save_close_button.BackColor = Color.Orange
        self.worksheet_save_close_button.Click += self.worksheet_save_close_clicked
        worksheet_group.Controls.Add(self.worksheet_save_close_button)
        
        self.delete_results_button = Button()
        self.delete_results_button.Text = "Delete Results"
        self.delete_results_button.Location = Point(115, 55)
        self.delete_results_button.Size = Size(95, 25)
        self.delete_results_button.BackColor = Color.LightCoral
        self.delete_results_button.Click += self.delete_results_clicked
        worksheet_group.Controls.Add(self.delete_results_button)
        
        # LIMS Export section
        lims_label = Label()
        lims_label.Text = "LIMS Export:"
        lims_label.Location = Point(10, 90)
        lims_label.Size = Size(80, 20)
        worksheet_group.Controls.Add(lims_label)
        
        self.lims_export_button = Button()
        self.lims_export_button.Text = "LIMS Export"
        self.lims_export_button.Location = Point(95, 85)
        self.lims_export_button.Size = Size(90, 30)  # Adjusted size
        self.lims_export_button.BackColor = Color.LightCyan
        self.lims_export_button.Click += self.lims_export_clicked
        worksheet_group.Controls.Add(self.lims_export_button)
        
        self.lims_location_button = Button()
        self.lims_location_button.Text = "Set LIMS Location"
        self.lims_location_button.Location = Point(195, 85)
        self.lims_location_button.Size = Size(100, 30)  # Adjusted position
        self.lims_location_button.BackColor = Color.LightPink
        self.lims_location_button.Click += self.lims_location_clicked
        worksheet_group.Controls.Add(self.lims_location_button)

        self.Controls.Add(worksheet_group)
    
    def create_sample_panel(self):
        """Create sample management panel"""
        sample_group = GroupBox()
        sample_group.Text = "Sample Management"
        sample_group.Location = Point(650, 10)  # Moved closer to left
        sample_group.Size = Size(280, 320)  # Reduced width and height
        
        # Sample list
        sample_label = Label()
        sample_label.Text = "Sample Queue:"
        sample_label.Location = Point(10, 25)
        sample_label.Size = Size(100, 20)
        sample_group.Controls.Add(sample_label)
        
        self.sample_listbox = ListBox()
        self.sample_listbox.Location = Point(10, 45)
        self.sample_listbox.Size = Size(260, 140)  # Adjusted size
        sample_group.Controls.Add(self.sample_listbox)
        
        # Sample selection for measurement
        selection_label = Label()
        selection_label.Text = "Sample Selection:"
        selection_label.Location = Point(10, 195)
        selection_label.Size = Size(120, 20)
        sample_group.Controls.Add(selection_label)
        
        self.select_for_measurement_button = Button()
        self.select_for_measurement_button.Text = "Select for Measurement"
        self.select_for_measurement_button.Location = Point(10, 215)
        self.select_for_measurement_button.Size = Size(125, 25)  # Adjusted size
        self.select_for_measurement_button.BackColor = Color.LightSalmon
        self.select_for_measurement_button.Click += self.select_for_measurement_clicked
        sample_group.Controls.Add(self.select_for_measurement_button)
        
        self.deselect_for_measurement_button = Button()
        self.deselect_for_measurement_button.Text = "Deselect"
        self.deselect_for_measurement_button.Location = Point(145, 215)
        self.deselect_for_measurement_button.Size = Size(75, 25)  # Adjusted size
        self.deselect_for_measurement_button.BackColor = Color.LightGray
        self.deselect_for_measurement_button.Click += self.deselect_for_measurement_clicked
        sample_group.Controls.Add(self.deselect_for_measurement_button)
        
        # Add sample controls
        add_sample_label = Label()
        add_sample_label.Text = "Add Sample:"
        add_sample_label.Location = Point(10, 250)
        add_sample_label.Size = Size(80, 20)
        sample_group.Controls.Add(add_sample_label)
        
        self.sample_name_textbox = TextBox()
        self.sample_name_textbox.Location = Point(10, 270)
        self.sample_name_textbox.Size = Size(110, 20)  # Adjusted size
        self.sample_name_textbox.Text = "Sample_001"
        sample_group.Controls.Add(self.sample_name_textbox)
        
        self.add_sample_button = Button()
        self.add_sample_button.Text = "Add"
        self.add_sample_button.Location = Point(130, 268)
        self.add_sample_button.Size = Size(40, 25)
        self.add_sample_button.Click += self.add_sample_clicked
        sample_group.Controls.Add(self.add_sample_button)
        
        self.clear_samples_button = Button()
        self.clear_samples_button.Text = "Clear All"
        self.clear_samples_button.Location = Point(180, 268)
        self.clear_samples_button.Size = Size(60, 25)
        self.clear_samples_button.Click += self.clear_samples_clicked
        sample_group.Controls.Add(self.clear_samples_button)
        
        self.Controls.Add(sample_group)
    
    def create_advanced_panel(self):
        """Create advanced controls panel"""
        advanced_group = GroupBox()
        advanced_group.Text = "Advanced Controls"
        advanced_group.Location = Point(650, 340)  # Moved to align with sample panel
        advanced_group.Size = Size(280, 110)  # Adjusted width to match sample panel
        
        # Process samples button
        self.process_samples_button = Button()
        self.process_samples_button.Text = "Process All Samples"
        self.process_samples_button.Location = Point(10, 25)
        self.process_samples_button.Size = Size(125, 30)  # Adjusted size
        self.process_samples_button.BackColor = Color.LightGreen
        self.process_samples_button.Click += self.process_samples_clicked
        advanced_group.Controls.Add(self.process_samples_button)
        
        # Export results button
        self.export_button = Button()
        self.export_button.Text = "Export Results"
        self.export_button.Location = Point(145, 25)
        self.export_button.Size = Size(95, 30)  # Adjusted size
        self.export_button.Click += self.export_clicked
        advanced_group.Controls.Add(self.export_button)
        
        # Output location button
        self.output_location_button = Button()
        self.output_location_button.Text = "Set Output Location"
        self.output_location_button.Location = Point(10, 65)
        self.output_location_button.Size = Size(115, 25)  # Adjusted size
        self.output_location_button.BackColor = Color.LightYellow
        self.output_location_button.Click += self.output_location_clicked
        advanced_group.Controls.Add(self.output_location_button)
        
        # Load worksheet template button
        self.load_template_button = Button()
        self.load_template_button.Text = "Load Template/Worksheet"
        self.load_template_button.Location = Point(135, 65)
        self.load_template_button.Size = Size(125, 25)  # Adjusted size
        self.load_template_button.BackColor = Color.LightCyan
        self.load_template_button.Click += self.load_template_clicked
        advanced_group.Controls.Add(self.load_template_button)
        
        self.Controls.Add(advanced_group)
    
    def create_log_panel(self):
        """Create log/message panel"""
        log_group = GroupBox()
        log_group.Text = "Activity Log"
        log_group.Location = Point(10, 460)  # Moved up to fit in smaller window
        log_group.Size = Size(920, 250)  # Reduced size for optimized layout
        
        self.log_textbox = TextBox()
        self.log_textbox.Multiline = True
        self.log_textbox.ScrollBars = ScrollBars.Vertical
        self.log_textbox.Location = Point(10, 20)
        self.log_textbox.Size = Size(900, 200)  # Adjusted for new panel size
        self.log_textbox.ReadOnly = True
        log_group.Controls.Add(self.log_textbox)
        
        self.clear_log_button = Button()
        self.clear_log_button.Text = "Clear Log"
        self.clear_log_button.Location = Point(10, 225)
        self.clear_log_button.Size = Size(80, 25)
        self.clear_log_button.Click += self.clear_log_clicked
        log_group.Controls.Add(self.clear_log_button)
        
        # Debug button to list available methods
        self.debug_button = Button()
        self.debug_button.Text = "Debug Methods"
        self.debug_button.Location = Point(100, 225)
        self.debug_button.Size = Size(100, 25)
        self.debug_button.Click += self.debug_methods_clicked
        log_group.Controls.Add(self.debug_button)
        
        self.Controls.Add(log_group)
    
    def log_message(self, message):
        """Add message to log panel"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = "[{0}] {1}\r\n".format(timestamp, message)
        self.log_textbox.AppendText(log_entry)
        
        # Keep log size manageable
        if len(self.log_textbox.Text) > 10000:
            lines = self.log_textbox.Lines
            self.log_textbox.Lines = lines[-100:]  # Keep last 100 lines
    
    def enable_controls(self, enabled):
        """Enable/disable instrument control buttons"""
        # Basic controls
        self.plasma_on_button.Enabled = enabled
        self.plasma_off_button.Enabled = enabled
        self.pump_off_button.Enabled = enabled
        self.pump_slow_button.Enabled = enabled
        self.pump_fast_button.Enabled = enabled
        self.purge_on_button.Enabled = enabled
        self.purge_off_button.Enabled = enabled
        self.start_button.Enabled = enabled
        self.stop_button.Enabled = enabled
        self.show_ui_button.Enabled = enabled
        self.hide_ui_button.Enabled = enabled
        self.process_samples_button.Enabled = enabled
        self.export_button.Enabled = enabled
        
        # Worksheet controls
        self.worksheet_new_button.Enabled = enabled
        self.worksheet_open_button.Enabled = enabled
        self.worksheet_save_button.Enabled = enabled
        self.worksheet_save_close_button.Enabled = enabled
        self.delete_results_button.Enabled = enabled  # Fixed button name
        self.lims_export_button.Enabled = enabled
        
        # Sample controls
        self.select_for_measurement_button.Enabled = enabled  # Fixed button name
        
        # Always enabled controls
        self.output_location_button.Enabled = True
        self.load_template_button.Enabled = True
        self.lims_location_button.Enabled = True  # Fixed button name
    
    # Event handlers
    def connect_clicked(self, sender, e):
        """Handle connect button click"""
        try:
            host = self.host_textbox.Text
            port_text = self.port_textbox.Text
            
            # Validate port number
            try:
                port = int(port_text)
                if port < 1 or port > 65535:
                    raise ValueError("Port must be between 1 and 65535")
            except ValueError as ve:
                self.log_message("Invalid port number: {0}".format(port_text))
                MessageBox.Show("Invalid port number: {0}".format(str(ve)), "Input Error", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                return
            
            self.log_message("Connecting to {0}:{1}...".format(host, port))
            
            # Clean up any existing connection
            if self.client:
                try:
                    self.client.Dispose()
                except:
                    pass  # Ignore cleanup errors
            
            self.client = Automation()
            
            # Connect and check if connection was successful
            self.client.Connect(host, port)
            
            # Check connection state from the client
            if hasattr(self.client.Client, 'State') and hasattr(self.client.Client.State, 'Connected'):
                # For XdrSocketClient, check the State property
                connected = (self.client.Client.State.value__ == 1)  # Connected state
            else:
                # Fallback: assume connected if no exception was thrown
                connected = True
            
            if connected:
                self.connected = True
                self.connection_status.Text = "Connected"
                self.connection_status.ForeColor = Color.Green
                self.connect_button.Enabled = False
                self.disconnect_button.Enabled = True
                self.enable_controls(True)
                self.status_timer.Start()
                self.log_message("Successfully connected to instrument")
            else:
                self.log_message("Failed to connect to instrument")
                MessageBox.Show("Failed to connect to instrument", "Connection Error", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                
        except Exception as ex:
            error_msg = str(ex)
            if "SocketException" in error_msg or "connection" in error_msg.lower():
                self.log_message("Connection failed: Unable to reach instrument at {0}:{1}".format(host, port_text))
                MessageBox.Show("Cannot connect to instrument.\n\nPlease check:\n- MP Expert is running\n- Host/Port are correct\n- Network connectivity", "Connection Failed", 
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            else:
                self.log_message("Connection error: {0}".format(error_msg))
                MessageBox.Show("Connection error: {0}".format(error_msg), "Error", 
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def disconnect_clicked(self, sender, e):
        """Handle disconnect button click"""
        try:
            if self.client and self.connected:
                # Stop the status timer first to prevent further communication
                if self.status_timer:
                    self.status_timer.Stop()
                
                self.log_message("Disconnecting from instrument...")
                
                # Attempt graceful disconnect
                try:
                    self.client.Disconnect()
                except Exception as disconnect_ex:
                    # Log the socket error but continue with cleanup
                    self.log_message("Socket disconnect warning: {0}".format(str(disconnect_ex)))
                    # This is expected when the connection is already broken
                
                # Clean up resources
                try:
                    self.client.Dispose()
                except Exception as dispose_ex:
                    self.log_message("Resource cleanup warning: {0}".format(str(dispose_ex)))
                
                # Update UI state
                self.connected = False
                self.connection_status.Text = "Not Connected"
                self.connection_status.ForeColor = Color.Red
                self.connect_button.Enabled = True
                self.disconnect_button.Enabled = False
                self.enable_controls(False)
                self.log_message("Disconnected from instrument successfully")
                
        except Exception as ex:
            self.log_message("Disconnect error: {0}".format(str(ex)))
    
    def plasma_on_clicked(self, sender, e):
        """Handle plasma ignite button click"""
        if self.client and self.connected:
            try:
                self.client.PlasmaOn()
                self.log_message("Plasma ignition command sent")
            except Exception as ex:
                self.log_message("Plasma ignition error: {0}".format(str(ex)))
    
    def plasma_off_clicked(self, sender, e):
        """Handle plasma extinguish button click"""
        if self.client and self.connected:
            try:
                self.client.PlasmaOff()
                self.log_message("Plasma extinguish command sent")
            except Exception as ex:
                self.log_message("Plasma extinguish error: {0}".format(str(ex)))
    
    def pump_off_clicked(self, sender, e):
        """Handle pump off button click"""
        if self.client and self.connected:
            try:
                self.client.PumpOff()
                self.log_message("Pump turned off")
            except Exception as ex:
                self.log_message("Pump control error: {0}".format(str(ex)))
    
    def pump_slow_clicked(self, sender, e):
        """Handle pump slow button click"""
        if self.client and self.connected:
            try:
                self.client.PumpSlow()
                self.log_message("Pump set to slow speed")
            except Exception as ex:
                self.log_message("Pump control error: {0}".format(str(ex)))
    
    def pump_fast_clicked(self, sender, e):
        """Handle pump fast button click"""
        if self.client and self.connected:
            try:
                self.client.PumpFast()
                self.log_message("Pump set to fast speed")
            except Exception as ex:
                self.log_message("Pump control error: {0}".format(str(ex)))
    
    def purge_on_clicked(self, sender, e):
        """Handle purge on button click"""
        if self.client and self.connected:
            try:
                self.client.PurgeOn()
                self.log_message("N2 purge enabled")
            except Exception as ex:
                self.log_message("Purge control error: {0}".format(str(ex)))
    
    def purge_off_clicked(self, sender, e):
        """Handle purge off button click"""
        if self.client and self.connected:
            try:
                self.client.PurgeOff()
                self.log_message("N2 purge disabled")
            except Exception as ex:
                self.log_message("Purge control error: {0}".format(str(ex)))
    
    def start_clicked(self, sender, e):
        """Handle start measurement button click"""
        if self.client and self.connected:
            try:
                self.client.Start()
                self.log_message("Measurement started")
            except Exception as ex:
                self.log_message("Start measurement error: {0}".format(str(ex)))
    
    def stop_clicked(self, sender, e):
        """Handle stop measurement button click"""
        if self.client and self.connected:
            try:
                self.client.Stop()
                self.log_message("Measurement stopped")
            except Exception as ex:
                self.log_message("Stop measurement error: {0}".format(str(ex)))
    
    def add_sample_clicked(self, sender, e):
        """Handle add sample button click"""
        sample_name = self.sample_name_textbox.Text.strip()
        if sample_name:
            self.sample_listbox.Items.Add(sample_name)
            self.sample_name_textbox.Text = "Sample_{0:03d}".format(self.sample_listbox.Items.Count + 1)
            self.log_message("Added sample: {0}".format(sample_name))
    
    def clear_samples_clicked(self, sender, e):
        """Handle clear samples button click"""
        self.sample_listbox.Items.Clear()
        self.log_message("Sample queue cleared")
    
    def process_samples_clicked(self, sender, e):
        """Handle process all samples button click"""
        if self.client and self.connected:
            if self.sample_listbox.Items.Count == 0:
                MessageBox.Show("No samples in queue", "Warning", 
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            
            try:
                self.log_message("Processing {0} samples...".format(self.sample_listbox.Items.Count))
                
                for i in range(self.sample_listbox.Items.Count):
                    sample_name = self.sample_listbox.Items[i]
                    self.log_message("Processing sample: {0}".format(sample_name))
                    self.client.SelectSolution(sample_name, True)
                    time.sleep(0.5)  # Brief delay between selections
                
                self.client.Start()
                self.log_message("Batch measurement started")
                
            except Exception as ex:
                self.log_message("Batch processing error: {0}".format(str(ex)))
    
    def export_clicked(self, sender, e):
        """Handle export results button click"""
        if self.client and self.connected:
            try:
                import os
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filename = "results_{0}.csv".format(timestamp)
                
                # Use selected output path or current directory
                if self.output_path:
                    export_path = os.path.join(self.output_path, filename)
                else:
                    export_path = filename
                
                self.client.Export(export_path)
                self.log_message("Results exported successfully!")
                self.log_message("File saved to: {0}".format(export_path))
                MessageBox.Show("Results exported successfully!\n\nFile saved to:\n{0}".format(export_path), "Export Complete", 
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            except Exception as ex:
                self.log_message("Export error: {0}".format(str(ex)))
    
    def clear_log_clicked(self, sender, e):
        """Handle clear log button click"""
        self.log_textbox.Clear()
    
    def debug_methods_clicked(self, sender, e):
        """Handle debug methods button click - lists available client methods"""
        if self.client:
            try:
                self.log_message("=== Available Client Methods ===")
                methods = [method for method in dir(self.client) if not method.startswith('_')]
                methods.sort()
                
                # Show methods in groups
                ui_methods = [m for m in methods if 'ui' in m.lower() or 'show' in m.lower() or 'hide' in m.lower()]
                if ui_methods:
                    self.log_message("UI Methods: {0}".format(', '.join(ui_methods)))
                
                # Show all methods in chunks to avoid overwhelming the log
                chunk_size = 15
                for i in range(0, len(methods), chunk_size):
                    chunk = methods[i:i+chunk_size]
                    self.log_message("Methods {0}-{1}: {2}".format(i+1, min(i+chunk_size, len(methods)), ', '.join(chunk)))
                
                self.log_message("=== End Method List ===")
                
            except Exception as ex:
                self.log_message("Debug methods error: {0}".format(str(ex)))
        else:
            self.log_message("No client connected - connect first to see available methods")
    
    def show_ui_clicked(self, sender, e):
        """Handle show MP Expert UI button click"""
        if self.client and self.connected:
            try:
                # Check if the method exists first
                if hasattr(self.client, 'ShowUI'):
                    self.client.ShowUI()
                    self.log_message("MP Expert UI shown")
                elif hasattr(self.client, 'ShowUserInterface'):
                    self.client.ShowUserInterface()
                    self.log_message("MP Expert UI shown")
                elif hasattr(self.client, 'Show'):
                    self.client.Show()
                    self.log_message("MP Expert UI shown")
                else:
                    # List available methods for debugging
                    methods = [method for method in dir(self.client) if not method.startswith('_')]
                    self.log_message("ShowUI method not found. Available methods: {0}".format(', '.join(methods[:10])))
                    self.log_message("Please check the Automation SDK documentation for the correct method name")
            except Exception as ex:
                self.log_message("Show UI error: {0}".format(str(ex)))
    
    def hide_ui_clicked(self, sender, e):
        """Handle hide MP Expert UI button click"""
        if self.client and self.connected:
            try:
                # Check if the method exists first
                if hasattr(self.client, 'HideUI'):
                    self.client.HideUI()
                    self.log_message("MP Expert UI hidden")
                elif hasattr(self.client, 'HideUserInterface'):
                    self.client.HideUserInterface()
                    self.log_message("MP Expert UI hidden")
                elif hasattr(self.client, 'Hide'):
                    self.client.Hide()
                    self.log_message("MP Expert UI hidden")
                else:
                    # List available methods for debugging
                    methods = [method for method in dir(self.client) if not method.startswith('_')]
                    self.log_message("HideUI method not found. Available methods: {0}".format(', '.join(methods[:10])))
                    self.log_message("Please check the Automation SDK documentation for the correct method name")
            except Exception as ex:
                self.log_message("Hide UI error: {0}".format(str(ex)))
    
    def output_location_clicked(self, sender, e):
        """Handle output location selection button click"""
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = "Select Output Location for Results"
        if self.output_path:
            folder_dialog.SelectedPath = self.output_path
        
        if folder_dialog.ShowDialog() == DialogResult.OK:
            self.output_path = folder_dialog.SelectedPath
            self.log_message("Output location set to: {0}".format(self.output_path))
            MessageBox.Show("Output location set to:\n{0}".format(self.output_path), "Output Location Set", 
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
    
    def load_template_clicked(self, sender, e):
        """Handle load worksheet template button click"""
        from System.Windows.Forms import OpenFileDialog, DialogResult
        
        file_dialog = OpenFileDialog()
        file_dialog.Title = "Select Template or Worksheet"
        file_dialog.Filter = "MP Expert Templates (*.mpts)|*.mpts|MP Expert Worksheets (*.mpws)|*.mpws|All Files (*.*)|*.*"
        file_dialog.FilterIndex = 1
        
        if file_dialog.ShowDialog() == DialogResult.OK:
            self.template_path = file_dialog.FileName
            file_type = "Template" if self.template_path.lower().endswith('.mpts') else "Worksheet" if self.template_path.lower().endswith('.mpws') else "File"
            self.log_message("{0} selected: {1}".format(file_type, self.template_path))
            
            # Actually load the template/worksheet into MP Expert
            if self.client and self.connected:
                try:
                    self.log_message("Loading {0} into MP Expert...".format(file_type.lower()))
                    
                    # Try different possible methods to load the file
                    loaded = False
                    
                    # Method 1: Try LoadTemplate
                    if hasattr(self.client, 'LoadTemplate'):
                        self.client.LoadTemplate(self.template_path)
                        loaded = True
                        self.log_message("{0} loaded successfully using LoadTemplate".format(file_type))
                    
                    # Method 2: Try LoadWorksheet  
                    elif hasattr(self.client, 'LoadWorksheet'):
                        self.client.LoadWorksheet(self.template_path)
                        loaded = True
                        self.log_message("{0} loaded successfully using LoadWorksheet".format(file_type))
                    
                    # Method 3: Try LoadFile
                    elif hasattr(self.client, 'LoadFile'):
                        self.client.LoadFile(self.template_path)
                        loaded = True
                        self.log_message("{0} loaded successfully using LoadFile".format(file_type))
                    
                    # Method 4: Try OpenFile
                    elif hasattr(self.client, 'OpenFile'):
                        self.client.OpenFile(self.template_path)
                        loaded = True
                        self.log_message("{0} loaded successfully using OpenFile".format(file_type))
                    
                    # Method 5: Try Load
                    elif hasattr(self.client, 'Load'):
                        self.client.Load(self.template_path)
                        loaded = True
                        self.log_message("{0} loaded successfully using Load".format(file_type))
                    
                    if loaded:
                        MessageBox.Show("{0} loaded successfully into MP Expert!\n\nFile: {1}".format(file_type, self.template_path), "{0} Loaded".format(file_type), 
                                      MessageBoxButtons.OK, MessageBoxIcon.Information)
                        
                        # Try to detect samples in the loaded file
                        self.detect_worksheet_samples()
                        
                        # Show the MP Expert UI if it's hidden
                        try:
                            if hasattr(self.client, 'ShowUI'):
                                self.client.ShowUI()
                            elif hasattr(self.client, 'ShowUserInterface'):
                                self.client.ShowUserInterface()
                            elif hasattr(self.client, 'Show'):
                                self.client.Show()
                        except Exception as show_ex:
                            self.log_message("Note: Could not show UI automatically: {0}".format(str(show_ex)))
                    
                    else:
                        # No suitable method found
                        available_methods = [method for method in dir(self.client) if not method.startswith('_') and ('load' in method.lower() or 'open' in method.lower())]
                        self.log_message("No suitable load method found. Available load/open methods: {0}".format(', '.join(available_methods)))
                        MessageBox.Show("Could not load {0}.\n\nNo suitable load method found in the Automation SDK.\nAvailable methods: {1}".format(file_type.lower(), ', '.join(available_methods[:5])), "Load Failed", 
                                      MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        
                except Exception as ex:
                    self.log_message("{0} loading error: {1}".format(file_type, str(ex)))
                    MessageBox.Show("Error loading {0}:\n\n{1}".format(file_type.lower(), str(ex)), "Load Error", 
                                  MessageBoxButtons.OK, MessageBoxIcon.Error)
            else:
                # Not connected - just show file selection confirmation
                MessageBox.Show("{0} selected:\n{1}\n\nConnect to instrument to load the file into MP Expert.".format(file_type, self.template_path), "{0} Selected".format(file_type), 
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
    
    def update_status(self, sender, e):
        """Update instrument status display"""
        if self.client and self.connected:
            try:
                # Check for new responses
                if hasattr(self.client, 'Responses') and len(self.client.Responses) > 0:
                    self.status_listbox.Items.Clear()
                    
                    # Process all available responses
                    responses_to_process = list(self.client.Responses)  # Create a copy
                    del self.client.Responses[:]  # Clear the original list (Python 2.7 compatible)
                    
                    for response in responses_to_process:
                        if isinstance(response, dict):
                            for key, value in response.items():
                                status_item = "{0}: {1}".format(key, value)
                                self.status_listbox.Items.Add(status_item)
                        else:
                            # Handle non-dict responses
                            self.status_listbox.Items.Add(str(response))
                    
                    # Keep only recent status items
                    while self.status_listbox.Items.Count > 20:
                        self.status_listbox.Items.RemoveAt(0)
                        
            except Exception as ex:
                # Check if this is a socket exception
                if "SocketException" in str(ex) or "established connection was aborted" in str(ex):
                    # Connection has been lost, update UI accordingly
                    self.log_message("Connection lost to instrument")
                    self.connected = False
                    self.connection_status.Text = "Connection Lost"
                    self.connection_status.ForeColor = Color.Orange
                    self.connect_button.Enabled = True
                    self.disconnect_button.Enabled = False
                    self.enable_controls(False)
                    if self.status_timer:
                        self.status_timer.Stop()
                else:
                    self.log_message("Status update error: {0}".format(str(ex)))

    # New event handlers for worksheet management
    def worksheet_new_clicked(self, sender, e):
        """Create new worksheet from template"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'WorksheetNew'):
                self.client.WorksheetNew()
                self.log_message("New worksheet created from template")
                
                # Try to detect samples in the new worksheet
                self.detect_worksheet_samples()
                
            else:
                self.log_message("WorksheetNew method not available")
        except Exception as ex:
            self.log_message("Error creating new worksheet: {0}".format(str(ex)))
    
    def worksheet_open_clicked(self, sender, e):
        """Open existing worksheet"""
        if not self.connected:
            return
        try:
            from System.Windows.Forms import OpenFileDialog, DialogResult
            dialog = OpenFileDialog()
            dialog.Filter = "Worksheet Files (*.mpws)|*.mpws|All Files (*.*)|*.*"
            dialog.Title = "Open Worksheet"
            
            if dialog.ShowDialog() == DialogResult.OK:
                if hasattr(self.client, 'WorksheetOpen'):
                    self.client.WorksheetOpen(dialog.FileName)
                    self.log_message("Worksheet opened: {0}".format(dialog.FileName))
                    
                    # Try to detect samples in the opened worksheet
                    self.detect_worksheet_samples()
                    
                else:
                    self.log_message("WorksheetOpen method not available")
        except Exception as ex:
            self.log_message("Error opening worksheet: {0}".format(str(ex)))
    
    def worksheet_save_clicked(self, sender, e):
        """Save current worksheet"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'WorksheetSaveAs'):
                from System.Windows.Forms import SaveFileDialog, DialogResult
                dialog = SaveFileDialog()
                dialog.Filter = "Worksheet Files (*.mpws)|*.mpws|All Files (*.*)|*.*"
                dialog.Title = "Save Worksheet"
                
                if dialog.ShowDialog() == DialogResult.OK:
                    self.client.WorksheetSaveAs(dialog.FileName)
                    self.log_message("Worksheet saved: {0}".format(dialog.FileName))
            else:
                self.log_message("WorksheetSaveAs method not available")
        except Exception as ex:
            self.log_message("Error saving worksheet: {0}".format(str(ex)))
    
    def worksheet_save_close_clicked(self, sender, e):
        """Save and close current worksheet"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'WorksheetSaveClose'):
                self.client.WorksheetSaveClose()
                self.log_message("Worksheet saved and closed")
            else:
                self.log_message("WorksheetSaveClose method not available")
        except Exception as ex:
            self.log_message("Error saving and closing worksheet: {0}".format(str(ex)))
    
    def worksheet_delete_results_clicked(self, sender, e):
        """Delete results from current worksheet"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'DeleteResults'):
                self.client.DeleteResults()
                self.log_message("Results deleted from worksheet")
            else:
                self.log_message("DeleteResults method not available")
        except Exception as ex:
            self.log_message("Error deleting results: {0}".format(str(ex)))
    
    def worksheet_close_clicked(self, sender, e):
        """Close current worksheet"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'WorksheetClose'):
                self.client.WorksheetClose()
                self.log_message("Worksheet closed")
            else:
                self.log_message("WorksheetClose method not available")
        except Exception as ex:
            self.log_message("Error closing worksheet: {0}".format(str(ex)))
    
    def lims_export_clicked(self, sender, e):
        """Export data to LIMS format"""
        if not self.connected:
            return
        try:
            if self.lims_output_path and hasattr(self.client, 'Export'):
                self.client.Export(self.lims_output_path)
                self.log_message("Data exported to LIMS format: {0}".format(self.lims_output_path))
                MessageBox.Show("Data exported successfully!\n\nFile saved to:\n{0}".format(self.lims_output_path), 
                              "LIMS Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            else:
                self.log_message("LIMS output path not set or Export method not available")
                MessageBox.Show("Please set LIMS output path first", "Export Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
        except Exception as ex:
            self.log_message("Error exporting to LIMS: {0}".format(str(ex)))
    
    def lims_browse_clicked(self, sender, e):
        """Browse for LIMS output path"""
        try:
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            dialog = FolderBrowserDialog()
            dialog.Description = "Select LIMS Export Folder"
            
            if dialog.ShowDialog() == DialogResult.OK:
                self.lims_output_path = dialog.SelectedPath
                self.log_message("LIMS output path set to: {0}".format(self.lims_output_path))
        except Exception as ex:
            self.log_message("Error setting LIMS path: {0}".format(str(ex)))
    
    # Fixed button name reference
    def lims_location_clicked(self, sender, e):
        """Browse for LIMS output path - same as lims_browse_clicked"""
        self.lims_browse_clicked(sender, e)
    
    def select_for_measurement_clicked(self, sender, e):
        """Select sample for measurement"""
        if self.sample_listbox.SelectedIndex >= 0:
            selected_sample = self.sample_listbox.SelectedItem
            self.log_message("Selected sample for measurement: {0}".format(selected_sample))
            
            # If connected, try to select the solution
            if self.client and self.connected:
                try:
                    if hasattr(self.client, 'SelectSolution'):
                        self.client.SelectSolution(str(selected_sample))
                        self.log_message("Solution selected in instrument: {0}".format(selected_sample))
                    else:
                        self.log_message("SelectSolution method not available")
                except Exception as ex:
                    self.log_message("Error selecting solution: {0}".format(str(ex)))
        else:
            self.log_message("No sample selected in list")
    
    def deselect_for_measurement_clicked(self, sender, e):
        """Deselect sample for measurement"""
        if self.sample_listbox.SelectedIndex >= 0:
            self.sample_listbox.ClearSelected()
            self.log_message("Sample deselected")
        else:
            self.log_message("No sample was selected")
    
    def delete_results_clicked(self, sender, e):
        """Delete results from current worksheet - fixed button name reference"""
        self.worksheet_delete_results_clicked(sender, e)
    
    def detect_worksheet_samples(self):
        """Detect and display samples from the currently opened worksheet"""
        try:
            # Clear existing sample queue
            self.sample_listbox.Items.Clear()
            
            # Try different methods to get sample information
            samples_found = False
            
            # Method 1: Try GetSamples if available
            if hasattr(self.client, 'GetSamples'):
                try:
                    samples = self.client.GetSamples()
                    if samples:
                        for sample in samples:
                            self.sample_listbox.Items.Add(str(sample))
                        samples_found = True
                        self.log_message("Found {0} samples in worksheet using GetSamples".format(len(samples)))
                except Exception as ex:
                    self.log_message("GetSamples error: {0}".format(str(ex)))
            
            # Method 2: Try GetSampleList if available
            if not samples_found and hasattr(self.client, 'GetSampleList'):
                try:
                    sample_list = self.client.GetSampleList()
                    if sample_list:
                        if hasattr(sample_list, '__iter__'):
                            for sample in sample_list:
                                self.sample_listbox.Items.Add(str(sample))
                        else:
                            # If it's a single item or string
                            self.sample_listbox.Items.Add(str(sample_list))
                        samples_found = True
                        sample_count = len(sample_list) if hasattr(sample_list, '__len__') else 1
                        self.log_message("Found {0} samples in worksheet using GetSampleList".format(sample_count))
                except Exception as ex:
                    self.log_message("GetSampleList error: {0}".format(str(ex)))
            
            # Method 3: Try GetWorksheetInfo or similar methods
            if not samples_found and hasattr(self.client, 'GetWorksheetInfo'):
                try:
                    worksheet_info = self.client.GetWorksheetInfo()
                    if worksheet_info and hasattr(worksheet_info, 'Samples'):
                        for sample in worksheet_info.Samples:
                            self.sample_listbox.Items.Add(str(sample))
                        samples_found = True
                        self.log_message("Found {0} samples in worksheet using GetWorksheetInfo".format(len(worksheet_info.Samples)))
                except Exception as ex:
                    self.log_message("GetWorksheetInfo error: {0}".format(str(ex)))
            
            # Method 4: Try GetSampleNames if available
            if not samples_found and hasattr(self.client, 'GetSampleNames'):
                try:
                    sample_names = self.client.GetSampleNames()
                    if sample_names:
                        for name in sample_names:
                            self.sample_listbox.Items.Add(str(name))
                        samples_found = True
                        self.log_message("Found {0} samples in worksheet using GetSampleNames".format(len(sample_names)))
                except Exception as ex:
                    self.log_message("GetSampleNames error: {0}".format(str(ex)))
            
            # Method 5: Try to get sample count and generate generic names
            if not samples_found and hasattr(self.client, 'GetSampleCount'):
                try:
                    count = self.client.GetSampleCount()
                    if count and count > 0:
                        for i in range(int(count)):
                            sample_name = "Sample_{0:03d}".format(i + 1)
                            self.sample_listbox.Items.Add(sample_name)
                        samples_found = True
                        self.log_message("Found {0} samples in worksheet using GetSampleCount".format(count))
                except Exception as ex:
                    self.log_message("GetSampleCount error: {0}".format(str(ex)))
            
            # If no samples found, log available methods for debugging
            if not samples_found:
                available_methods = [method for method in dir(self.client) 
                                   if not method.startswith('_') and 
                                   ('sample' in method.lower() or 'worksheet' in method.lower())]
                if available_methods:
                    self.log_message("No sample detection method worked. Available sample/worksheet methods: {0}".format(', '.join(available_methods[:10])))
                else:
                    self.log_message("No sample detection methods available in SDK")
                
                # Add a default sample as placeholder
                self.sample_listbox.Items.Add("Sample_001")
                self.log_message("Added default sample placeholder - worksheet samples not auto-detected")
            
        except Exception as ex:
            self.log_message("Error detecting worksheet samples: {0}".format(str(ex)))
            # Add a default sample as fallback
            self.sample_listbox.Items.Add("Sample_001")
            self.log_message("Added default sample due to detection error")
    
    def select_solution_clicked(self, sender, e):
        """Select solution for measurement"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'SelectSolution'):
                # You could add a dialog here to select which solution
                self.client.SelectSolution()
                self.log_message("Solution selected for measurement")
            else:
                self.log_message("SelectSolution method not available")
        except Exception as ex:
            self.log_message("Error selecting solution: {0}".format(str(ex)))
    
    def get_version_clicked(self, sender, e):
        """Get software version information"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'GetVersion'):
                version = self.client.GetVersion()
                self.log_message("Software version: {0}".format(str(version)))
            else:
                self.log_message("GetVersion method not available")
        except Exception as ex:
            self.log_message("Error getting version: {0}".format(str(ex)))
    
    def get_status_clicked(self, sender, e):
        """Get detailed instrument status"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'GetStatus'):
                status = self.client.GetStatus()
                self.log_message("Instrument status: {0}".format(str(status)))
            else:
                self.log_message("GetStatus method not available")
        except Exception as ex:
            self.log_message("Error getting status: {0}".format(str(ex)))
    
    def ready_clicked(self, sender, e):
        """Set instrument to ready state"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'Ready'):
                self.client.Ready()
                self.log_message("Instrument set to ready state")
            else:
                self.log_message("Ready method not available")
        except Exception as ex:
            self.log_message("Error setting ready state: {0}".format(str(ex)))
    
    def standby_clicked(self, sender, e):
        """Set instrument to standby state"""
        if not self.connected:
            return
        try:
            if hasattr(self.client, 'Standby'):
                self.client.Standby()
                self.log_message("Instrument set to standby state")
            else:
                self.log_message("Standby method not available")
        except Exception as ex:
            self.log_message("Error setting standby state: {0}".format(str(ex)))

def main():
    """Main entry point for GUI application"""
    Application.EnableVisualStyles()
    Application.SetCompatibleTextRenderingDefault(False)
    
    # Create and run the GUI
    app = InstrumentControlGUI()
    Application.Run(app)

if __name__ == '__main__':
    main()
