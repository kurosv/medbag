import os


# Class; GetDesktopPath
#     *arg = none
#     get your desktop path

class GetDesktopPath:
    def __init__(self):
        self.desktop_path = None
        self.get_desktop_path()

    def get_desktop_path(self):
        self.desktop_path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\Desktop"


if __name__ == '__main__':
    dp = GetDesktopPath()
    print(dp.desktop_path)


# Class; SetDirPath
#   *arg = dir name
#   super: GetDesktopPath
#   get your desktop path
#   and set your unique desktop dir path

class SetDirPath(GetDesktopPath):
    def __init__(self, dir_name):
        super().__init__()
        self.dir_path = self.desktop_path
        self.set_path(dir_name)

    def set_path(self, dir_name):
        self.dir_path += "\\" + dir_name


if __name__ == '__main__':
    d_name = "test"
    sp = SetDirPath(d_name)
    print(sp.desktop_path)
    print(sp.dir_path)
