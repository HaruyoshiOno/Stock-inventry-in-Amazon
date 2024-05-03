import winreg

install_thunderbird = False

def check_thunderbird():
    try:
        global install_thunderbird
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Mozilla\Thunderbird")
        winreg.CloseKey(key)
        install_thunderbird = True
        print("Thunderbird is installed.")

    except FileNotFoundError:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Thunderbird")
            winreg.CloseKey(key)
            install_thunderbird = True
            print("Thunderbird is installed.")

        except FileNotFoundError:
            print("Thunderbird is not installed.")