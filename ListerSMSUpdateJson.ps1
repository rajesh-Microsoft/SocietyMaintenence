import asyncio
import requests
from winrt.windows.ui.notifications.management import UserNotificationListener, UserNotificationListenerAccessStatus

API_URL = "https://your-api-endpoint.com/sms"  # Replace with your REST API endpoint

async def listen_sms():
    listener = UserNotificationListener.get_current()

    # Request access to notifications
    access = await listener.request_access_async()
    if access != UserNotificationListenerAccessStatus.ALLOWED:
        print("Access denied to notifications. Exiting.")
        return

    seen_notifications = set()

    print("Listening for new SMS from Phone Link...")
    while True:
        notifications = await listener.get_notifications_async()
        for n in notifications:
            if "Your Phone" in n.app_display_name:  # Filter Phone Link notifications
                try:
                    text_elements = n.notification.visual.get_binding("ToastGeneric").text_elements
                    if len(text_elements) >= 2:
                        sender = text_elements[0].text   # Usually the sender number or name
                        body = text_elements[1].text     # SMS body
                        
                        # Avoid duplicates
                        if n.id not in seen_notifications:
                            seen_notifications.add(n.id)

                            # Prepare payload
                            payload = {"sender": sender, "message": body}

                            # Send to REST API
                            try:
                                response = requests.post(API_URL, json=payload)
                                print(f"Sent SMS to API: {payload}")
                                print(f"API response: {response.status_code} - {response.text}")
                            except Exception as e:
                                print("Error sending to API:", e)
                except Exception as e:
                    print("Error reading notification:", e)

        await asyncio.sleep(2)  # Check every 2 seconds

if __name__ == "__main__":
    asyncio.run(listen_sms())
