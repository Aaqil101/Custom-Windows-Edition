# -------------------------------
# Cleanup file/folder
# -------------------------------
Remove-Item -Path "C:\Recovery.txt" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\SetupFiles\" -Recurse -Force -ErrorAction SilentlyContinue
