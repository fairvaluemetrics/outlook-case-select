Replace these placeholders with real PNG icons before shipping:
  icon-16.png   16x16
  icon-32.png   32x32
  icon-64.png   64x64
  icon-80.png   80x80
  icon-128.png  128x128

The manifest references them from /assets/. Outlook will fall back to a
default icon if a referenced file is missing, but sideload validation
will warn about it.
