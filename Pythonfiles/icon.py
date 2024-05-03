import tkinter as tk

icon_data = """R0lGODlhIAAgALMOAN66nN7PvYyqUoxVcwAAAGMwEK26nL2KUoyKjM6qczEQAHNF
            MWNlYzEwMf///wAAACH/C1hNUCBEYXRhWE1QPD94cGFja2V0IGJlZ2luPSLvu78i
            IGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxu
            czp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgOS4x
            LWMwMDEgNzkuMTQ2Mjg5OTc3NywgMjAyMy8wNi8yNS0yMzo1NzoxNCAgICAgICAg
            Ij4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAy
            LzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9
            IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxu
            czp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6
            c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJj
            ZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIDI0LjcgKFdp
            bmRvd3MpIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOkEwRjQ4OTJCREQ3QjEx
            RUVCRUU3Q0RGMDgxRDFDMTI0IiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOkEw
            RjQ4OTJDREQ3QjExRUVCRUU3Q0RGMDgxRDFDMTI0Ij4gPHhtcE1NOkRlcml2ZWRG
            cm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6QTBGNDg5MjlERDdCMTFFRUJF
            RTdDREYwODFEMUMxMjQiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6QTBGNDg5
            MkFERDdCMTFFRUJFRTdDREYwODFEMUMxMjQiLz4gPC9yZGY6RGVzY3JpcHRpb24+
            IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz4B//79
            /Pv6+fj39vX08/Lx8O/u7ezr6uno5+bl5OPi4eDf3t3c29rZ2NfW1dTT0tHQz87N
            zMvKycjHxsXEw8LBwL++vby7urm4t7a1tLOysbCvrq2sq6qpqKempaSjoqGgn56d
            nJuamZiXlpWUk5KRkI+OjYyLiomIh4aFhIOCgYB/fn18e3p5eHd2dXRzcnFwb25t
            bGtqaWhnZmVkY2JhYF9eXVxbWllYV1ZVVFNSUVBPTk1MS0pJSEdGRURDQkFAPz49
            PDs6OTg3NjU0MzIxMC8uLSwrKikoJyYlJCMiISAfHh0cGxoZGBcWFRQTEhEQDw4N
            DAsKCQgHBgUEAwIBAAAh+QQBAAAOACwAAAAAIAAgAAAE/9DJSau9OFO0UNBg1TBK
            gTRGGB5d4yJAo2bJUrgKwrjLbCWJQ4Kx0xFdPgliCAwuXLsGgOFrHA474XWxyD13
            VBWj2WQ9G4XdgtGbGd4GhHDtcg3ZScdYe6wTbwVJBkEsDScNClBQeQ5BOA0BiS9E
            jEEIBnUICgYkLh8+AQd1o4lqKEkJmpI3jySUbl0KBIiSI4g2bRoGAEFRCr+jfgC5
            FwcAxxy1DU8Fvww6iCk0QkuIdmejT2cZC0BXdcDKUjgDYRZx3qyaLpi4J0RsxBRx
            Cwc6OpJRVjzY0hZ7a7qAsRZjhI0Ya/xVYGDgiYFf726sM3jgTYcMAEwQEehCHRcb
            RDwCnMJgjNMRG1YUrEmDpkABTiB4rWEg4GONAgA4dHGpEIOBAmsSGLTjadmTPLsu
            vTA0gsEnRheeQp1aIQIAOw==
        """

def get_photo_image4icon():
    return tk.PhotoImage(data=icon_data)