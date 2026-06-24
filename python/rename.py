import os
import sys
import shutil
import time

terminal_width = shutil.get_terminal_size().columns

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def msg(text=None):
    if not text:
        raise Exception('Message cannot be empty')
    print(text)
    print('-' * terminal_width)
    time.sleep(1)
    clear()

def pause_and_exit():
    input("Press Enter to exit . . . ")
    exit()

def rename_files(directory):
    files = sorted([f for f in os.listdir(directory) if f.endswith(".sor")])
    grouped_files = {}
    
    for file in files:
        parts = file.split("_")
        fiber_number = parts[0][5:]
        if fiber_number not in grouped_files:
            grouped_files[fiber_number] = []
        grouped_files[fiber_number].append(file)
    
    for fiber_number, file_list in grouped_files.items():
        print(f"Fiber{fiber_number} rename to: ", end="")
        new_number = input().strip()
        
        if new_number:
            for file in file_list:
                parts = file.split("_")
                new_name = f"Fiber{new_number}_{parts[1]}"
                old_path = os.path.join(directory, file)
                new_path = os.path.join(directory, new_name)
                os.rename(old_path, new_path)
                print('-' * terminal_width)
                msg(f"{file} renamed to {new_name}")
        else:
            print('-' * terminal_width)
            msg(f"Fiber{fiber_number} files not renamed")

if __name__ == "__main__":
    try:
        folder_path = sys.argv[1] if len(sys.argv) > 1 else input("Enter the folder path: ").strip()
        if os.path.exists(folder_path):
            rename_files(folder_path)
        else:
            print("Invalid folder path!")
    except Exception as e:
        print(f"An error occurred: {e}")
        pause_and_exit()