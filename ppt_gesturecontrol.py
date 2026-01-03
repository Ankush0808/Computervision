import cv2
import pyautogui
import time
from cvzone.HandTrackingModule import HandDetector
import pygetwindow as gw

# --- Try to activate PowerPoint window ---
def activate_powerpoint():
    try:
        windows = gw.getWindowsWithTitle('PowerPoint')
        if windows:
            windows[0].activate()
            print("✅ PowerPoint window activated.")
            time.sleep(1)
        else:
            print("⚠️ PowerPoint window not found. Please open your slideshow first (F5).")
    except Exception as e:
        print("Error activating PowerPoint:", e)

# --- Initialize webcam and hand detector ---
cap = cv2.VideoCapture(0)
detector = HandDetector(detectionCon=0.8, maxHands=1)

activate_powerpoint()

last_action_time = 0
cooldown = 1.0  # seconds

print("\n🤖 Gesture control started. Use:")
print("✌️  Two fingers  → Play/Pause video")
print("☝️  One finger   → Next slide")
print("🤘  Pinky only   → Previous slide")
print("Press 'q' to quit.\n")

while True:
    success, img = cap.read()
    img = cv2.flip(img, 1)
    hands, img = detector.findHands(img)

    if hands:
        hand = hands[0]
        fingers = detector.fingersUp(hand)
        current_time = time.time()

        # Gesture 1: Index + Middle → Play/Pause
        if fingers == [0, 1, 1, 0, 0] and current_time - last_action_time > cooldown:
            pyautogui.press('space')
            print("▶️ Play/Pause Video")
            last_action_time = current_time

        # Gesture 2: Index only → Next Slide
        elif fingers == [0, 1, 0, 0, 0] and current_time - last_action_time > cooldown:
            pyautogui.press('right')
            print("⏭️ Next Slide")
            last_action_time = current_time

        # Gesture 3: Pinky only → Previous Slide
        elif fingers == [0, 0, 0, 0, 1] and current_time - last_action_time > cooldown:
            pyautogui.press('left')
            print("⏮️ Previous Slide")
            last_action_time = current_time

    cv2.imshow("PowerPoint Gesture Controller", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
