
## Requirement Document: PowerPoint Control with Hand Gestures

### 1. Overview
The purpose of this project is to control a PowerPoint presentation using hand gestures captured through a webcam. The code allows users to advance to the next slide, go back to the previous slide, return to the first slide, and navigate to the last slide. Additionally, it includes features such as zooming in and out, activating pen mode for annotation, and drawing on slides.

### 2. Features
- Control PowerPoint slides using hand gestures detected from a webcam.
- Advance to the next slide by raising all fingers.
- Go back to the previous slide by raising the index and middle fingers.
- Return to the first slide by raising the index finger.
- Navigate to the last slide by lowering all fingers except the thumb.
- Zoom in and out of slides using specific hand gestures.
- Activate pen mode for annotation by raising the index finger and draw on slides using the hand's movement.

### 3. Dependencies
- Python 3.x
- OpenCV (cv2)
- PyAutoGUI
- win32com.client
- cvzone (for HandTrackingModule)
- aspose.slides
- aspose.pydrawing

### 4. Usage
1. Install the required Python packages.
2. Connect a webcam to your computer.
3. Run the Python script.
4. Present your hand gestures in front of the webcam to control the PowerPoint presentation.

### 5. Configuration
- Adjust the `gestureThreshold` and `delay` variables to fine-tune gesture detection sensitivity and button press delay.
- Specify the path to your PowerPoint presentation file in the `Presentation.Presentations.Open()` method.

### 6. Future Improvements
- Implement additional gestures for more commands, such as pausing/resuming the presentation or skipping to specific slides.
- Enhance the drawing functionality by adding options for different colors, line thicknesses, and erasing.
- Improve the robustness of gesture detection and accuracy of hand tracking.

### 7. Known Issues
- The code may not work properly with certain webcam configurations or lighting conditions.
- Hand gesture recognition accuracy may vary depending on factors such as hand size, orientation, and background clutter.

