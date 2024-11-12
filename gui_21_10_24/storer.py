import sys
import json
import ssl
import paho.mqtt.client as mqtt
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QDialogButtonBox, QMessageBox, QWidget, QPushButton

# Mock MQTT Handler for demonstration (replace with your actual MQTT handler)
class MqttHandler(QThread):
    mqtt_message_signal = pyqtSignal(str, str)  # Signal to pass the topic and payload

    def __init__(self, broker, port, username, password):
        super().__init__()
        self.broker = broker
        self.port = port
        self.username = username
        self.password = password
        self.client = mqtt.Client(client_id="unique_client_id")
        self.client.username_pw_set(self.username, self.password)
        self.client.tls_set(cert_reqs=ssl.CERT_NONE)
        self.client.tls_insecure_set(True)
        self.client.on_connect = self.on_connect
        self.client.on_message = self.on_message

    def run(self):
        self.client.connect(self.broker, self.port, 60)
        self.client.loop_forever()

    def on_connect(self, client, userdata, flags, rc):
        if rc == 0:
            print("Connected to MQTT Broker!")
            client.subscribe("#", qos=1)
        else:
            print(f"Failed to connect, return code {rc}")

    def on_message(self, client, userdata, msg):
        topic = msg.topic
        try:
            payload = msg.payload.decode('utf-8')  # Try decoding with UTF-8
        except UnicodeDecodeError:
            # Handle non-UTF-8 payloads or binary data
            payload = f"Binary data received: {msg.payload}"
        self.mqtt_message_signal.emit(topic, payload)


# Worker class for the background test process
class TestWorker(QThread):
    update_message = pyqtSignal(str)  # Signal to send update messages to the UI
    finished = pyqtSignal()  # Signal to notify when the worker has finished
    
    def __init__(self, parent, selected_rows, data, duration, count, selected_topic, mqtt_handler):
        super().__init__(parent)
        self.selected_rows = selected_rows
        self.data = data
        self.duration = duration
        self.count = count
        self.selected_topic = selected_topic
        self.mqtt_handler = mqtt_handler
        
    def run(self):
        # Perform the long-running task (publishing test to devices)
        self.selected_count = sum(1 for row_data in self.data if row_data[5])
        
        # Construct the message
        message = (
            f"<div style='text-align: center;'>"
            f"<b>Publishing LED glow test, '{self.selected_topic}' for the selected {self.selected_count} devices.</b><br>"
            f"<b>Blink duration: {self.duration} seconds, Blink count: {self.count}</b><br>"
            "Note: Go to settings to change any of the parameters."
            "</div>"
        )
        # Emit the message to update the UI
        self.update_message.emit(message)
        
        # Loop through each selected row and publish the message
        for row_index in self.selected_rows:
            mac_address = self.data[row_index][0]
            payload = f"LG {self.duration},{self.count}"
            self.publish_message(mac_address, payload)
        
        # Finish the task
        self.finished.emit()

    def publish_message(self, mac_address, payload):
        # Construct the message and publish
        topic = self.selected_topic
        message = {
            "devID": mac_address.upper(),
            "data": payload
        }
        self.mqtt_handler.client.publish(topic, json.dumps(message), qos=1)


# Main UI class
class YourClass(QWidget):
    def __init__(self):
        super().__init__()
        
        # MQTT Broker details
        self.MQTT_BROKER = "103.162.246.109"  # Replace with actual broker
        self.MQTT_PORT = 8883
        self.MQTT_USERNAME = "mqtt"
        self.MQTT_PASSWORD = "mqtt2022"
        
        # Initialize the MQTT handler
        self.mqtt_handler = MqttHandler(self.MQTT_BROKER, self.MQTT_PORT, self.MQTT_USERNAME, self.MQTT_PASSWORD)
        
        # Connect signal to handle incoming messages
        self.mqtt_handler.mqtt_message_signal.connect(self.on_mqtt_message)

        # Start the MQTT handler thread
        self.mqtt_handler.start()

        # Test parameters (you can modify these)
        self.selected_topic = 'your/topic'  # Replace with actual topic
        self.duration = 5  # Example duration
        self.count = 3  # Example blink count
        
        # Sample data (MAC addresses and selected status)
        self.data = [
            ['00:11:22:33:44:55', None, None, None, None, True],
            ['66:77:88:99:00:11', None, None, None, None, True],
            ['AA:BB:CC:DD:EE:FF', None, None, None, None, False],
        ]
        
        # Create a button to trigger the test
        self.run_test_button = QPushButton("Run Test", self)
        self.run_test_button.clicked.connect(self.run_test_action_triggered)
        
        layout = QVBoxLayout()
        layout.addWidget(self.run_test_button)
        self.setLayout(layout)
        
        self.setWindowTitle("Test Runner")
        self.setGeometry(300, 300, 400, 200)

    def run_test_action_triggered(self):
        selected_rows = self.get_selected_rows()
        if not selected_rows:
            error_dialog = QMessageBox(self)
            error_dialog.setIcon(QMessageBox.Warning)
            error_dialog.setText("Please select at least one row.")
            error_dialog.setWindowTitle("Error")
            error_dialog.exec_()
            return
        
        # Confirmation dialog before running the test
        custom_dialog = QDialog(self)
        custom_dialog.setWindowTitle("Run Test")
        custom_dialog.setFixedSize(550, 150)

        layout = QVBoxLayout()
        message_label = QLabel("Are you sure you want to run the test for the selected devices?")
        message_label.setWordWrap(True)
        layout.addWidget(message_label)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, custom_dialog)
        layout.addWidget(button_box)

        custom_dialog.setLayout(layout)

        # Connect buttons
        button_box.accepted.connect(custom_dialog.accept)
        button_box.rejected.connect(custom_dialog.reject)

        # Show the confirmation dialog and wait for the result
        reply = custom_dialog.exec_()

        if reply == QDialog.Accepted:
            # Start the background worker
            self.test_worker = TestWorker(
                self,
                selected_rows=selected_rows,
                data=self.data,
                duration=self.duration,
                count=self.count,
                selected_topic=self.selected_topic,
                mqtt_handler=self.mqtt_handler
            )
            
            # Connect signals
            self.test_worker.update_message.connect(self.on_test_update_message)
            self.test_worker.finished.connect(self.on_test_finished)
            
            # Start the worker thread
            self.test_worker.start()

    def on_mqtt_message(self, topic, payload):
        # Handle incoming MQTT messages (this function can update the UI with new messages)
        print(f"Received message from topic {topic}: {payload}")

    def on_test_update_message(self, message):
        # Update UI with messages from the background worker
        print(message)  # For debugging, you can display it in a QLabel or QTextEdit
        
    def on_test_finished(self):
        # Handle post-test actions, e.g., show a message that the test is complete
        print("Test finished.")

    def get_selected_rows(self):
        # Return the indices of rows that are selected (where the 6th column is True)
        selected_rows = [index for index, row in enumerate(self.data) if row[5]]
        return selected_rows


# Main execution
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = YourClass()
    window.show()
    sys.exit(app.exec_())
