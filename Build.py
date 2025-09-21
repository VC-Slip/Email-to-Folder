import subprocess
import os
import sys
import shutil

# ------------------------
# CONFIGURATION
# ------------------------
extra_files = [
    ("config.json", "."),  # single file
    ("images/", "images")  # folder
]

# Dry run mode: True = simulate, False = actually build exe
DRY_RUN = True

# Initialize main_script to None
main_script = None
# Optional: manually set main script name
# main_script = "your_script.py"

# ------------------------
# Auto-detect main script
# ------------------------
if main_script is None:
    current_folder = os.getcwd()
    py_files = [
        f for f in os.listdir(current_folder)
        if f.endswith(".py") and f.lower() != "build.py" and not f.startswith(("_", "test"))
    ]

    if not py_files:
        print("❌ No suitable Python script found to build!")
        sys.exit(1)

    # Heuristic: pick the first file that isn’t the build script
    main_script = py_files[0]

print(f"📝 Main script: {main_script}")

# ------------------------
# Clean old builds
# ------------------------
print("\n🧹 Dry run: checking old build artifacts...")
for folder in ["build", "dist", f"{os.path.splitext(main_script)[0]}.spec"]:
    if os.path.exists(folder):
        if os.path.isdir(folder):
            print(f" - Would remove folder: {folder}")
            if not DRY_RUN:
                shutil.rmtree(folder)
        else:
            print(f" - Would remove file: {folder}")
            if not DRY_RUN:
                os.remove(folder)
    else:
        print(f" - No existing {folder} to remove")

# ------------------------
# Prepare extra files for PyInstaller
# ------------------------
print("\n📦 Dry run: preparing extra files to include...")
add_data_args = []
for src, dest in extra_files:
    sep = ";" if sys.platform.startswith("win") else ":"
    add_data_args.extend(["--add-data", f"{src}{sep}{dest}"])
    # Verbose output
    if os.path.exists(src):
        print(f" - Would add: {src} -> {dest}")
    else:
        print(f" - Extra file/folder not found (would include if it existed): {src} -> {dest}")

# ------------------------
# Build command
# ------------------------
cmd = ["pyinstaller", "--onefile"] + add_data_args + [main_script]

print("\n🔹 PyInstaller command prepared:")
print(" ".join(cmd))

# ------------------------
# Execute build (if not dry run)
# ------------------------
if DRY_RUN:
    print("\n💡 Dry run enabled — no exe created. All automation steps simulated.")
else:
    print("\n🚀 Building EXE...")
    subprocess.run(cmd)
    print("✅ Build complete! Check the dist/ folder.")
