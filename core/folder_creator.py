import os


class FolderCreator:
    @staticmethod
    def create_photo_folders(
        photo_root: str,
        day_date_str: str,
        new_shipyard_tasks: list,
        new_engine_room_tasks: list,
    ) -> list:
        created_folders = []

        # Create main day folder
        photo_day_folder = os.path.join(photo_root, day_date_str)
        os.makedirs(photo_day_folder, exist_ok=True)
        created_folders.append(photo_day_folder)

        # Shipyard task folders (no date prefix)
        for _, line in new_shipyard_tasks:
            if line and line != "nan":
                folder_path = os.path.join(photo_day_folder, line.strip())
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)
                    created_folders.append(folder_path)

        # Engine room task folders (with date prefix)
        for _, line in new_engine_room_tasks:
            if line and line != "nan":
                folder_name = f"{day_date_str}_{line.strip()}"
                folder_path = os.path.join(photo_day_folder, folder_name)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)
                    created_folders.append(folder_path)

        return created_folders
