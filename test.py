import mouse
import json

points = json.load(open("ui_points.json"))
x, y = points["camera_tab"]["x"], points["camera_tab"]["y"]
print(f"Moving to ({x}, {y})")
mouse.move(x, y, absolute=True, duration=0.5)
