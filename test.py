import mouse
import time
import json

# Load your points
points = json.load(open("ui_points.json"))

print("This will click your recorded points in 5 seconds...")
print("Make sure Visualize Import Settings dialog is open!")
time.sleep(5)

# Click Import OK
x, y = points["import_ok_btn"]["x"], points["import_ok_btn"]["y"]
print(f"Clicking import_ok_btn at ({x}, {y})")
mouse.move(x, y, absolute=True, duration=0.5)
time.sleep(1)
mouse.click()
print("Clicked!")
