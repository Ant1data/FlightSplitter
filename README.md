# FlightSplitter - Air France Flight Rotation

This project automatically generates Excel files of Air France flight rotations from a source file.
GUI developed in Python with Tkinter, progress bar, and animation.

## How to Use

1. Double-click `lancer_rotation.bat` to launch the interface (no black console).
2. Choose the source Excel file.
3. Choose the output folder.
4. Click Start Processing.
5. The Excel files will be generated and the output folder will open automatically.

## Dependencies

- Python 3.x (including Tkinter)
- pandas
- openpyxl

To create the required Python environment
```bash
conda create -n mini-flightsplitter pandas openpyxl


