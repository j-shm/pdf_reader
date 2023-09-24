import os
import shutil
import subprocess
import zipfile

# Full paths to source files and folders
source_folder = os.getcwd()
destination_folder = os.getcwd()+"/packed_program/"

# Step 1: Copy files
def copy_files():
    try:
        # Create the destination folder if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        # Copy program.py, run.bat, and poppler into the destination folder
        files_to_copy = ["program.py", "run.bat", os.path.join(source_folder, "poppler")]
        for file in files_to_copy:
            if os.path.isfile(file):
                shutil.copy(file, os.path.join(destination_folder, os.path.basename(file)))
            elif os.path.isdir(file):
                shutil.copytree(file, os.path.join(destination_folder, os.path.basename(file)))

        return destination_folder
    except Exception as e:
        print(f"Error copying files: {e}")
        return None

# Step 2: Generate requirements.txt using pipreqs
def generate_requirements_txt(destination_folder):
    try:
        subprocess.call(["pipreqs", destination_folder])
    except Exception as e:
        print(f"Error generating requirements.txt: {e}")

# Step 3: Create a ZIP file of the packed folder
def create_zip(destination_folder):
    try:
        with zipfile.ZipFile("packed_program.zip", "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(destination_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, destination_folder)
                    zipf.write(file_path, arcname)

        print("Packing completed. ZIP file created.")
    except Exception as e:
        print(f"Error creating ZIP file: {e}")


if __name__ == "__main__":
    # Step 1: Copy files
    copied_folder = copy_files()
    
    if copied_folder:
        # Step 2: Generate requirements.txt
        generate_requirements_txt(copied_folder)

        # Step 3: Create a ZIP file
        create_zip(copied_folder)

        # Step 4: Delete the copied folder
        shutil.rmtree(copied_folder)