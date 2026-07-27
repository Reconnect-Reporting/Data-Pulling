from pathlib import Path
import shutil


def move_stimes_downloads_to_sound_times():
    """
    Move all downloaded Stimes files from local Downloads folder to:

    C:\\Users\\<username>\\OneDrive - Reconnect Community Health Services (1)\\data\\Raw Data\\Sound Times

    Auto-detects the Windows username using Path.home().

    If a file with the same name already exists in the destination folder,
    it will be replaced.
    """

    user_home = Path.home()

    downloads_folder = user_home / "Downloads"

    destination_folder = (
        user_home
        / "OneDrive - Reconnect Community Health Services (1)"
        / "data"
        / "Raw Data"
        / "Sound Times"
    )

    destination_folder.mkdir(parents=True, exist_ok=True)

    files_to_move = [
        file for file in downloads_folder.iterdir()
        if file.is_file()
        and "_Stimes" in file.stem
        and not file.name.endswith(".crdownload")
    ]

    if not files_to_move:
        print("No _Stimes files found in Downloads folder.")
        return

    for src_file in files_to_move:
        dst_file = destination_folder / src_file.name

        # Replace file if same name already exists
        if dst_file.exists():
            dst_file.unlink()
            print(f"Existing file replaced: {dst_file.name}")

        shutil.move(str(src_file), str(dst_file))

        print("Moved:")
        print(f"  From: {src_file}")
        print(f"  To:   {dst_file}")

    print("All Stimes downloaded files moved successfully.")


def main():
    move_stimes_downloads_to_sound_times()


if __name__ == "__main__":
    main()