import os
import getPath


# Class; SetDir
#     *arg = dir name
#     make dir to your Desktop

class SetDir:
    def __init__(self, dir_name):
        gp = getPath.SetDirPath(dir_name)
        self.path = gp.dir_path
        self.set_dir()

    def set_dir(self):
        if not os.path.isdir(self.path):
            os.makedirs(self.path)


if __name__ == '__main__':
    d_name = "test"
    sd = SetDir(d_name)
