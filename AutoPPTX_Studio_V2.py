"""AutoPPTX Studio V2 — chạy: python AutoPPTX_Studio_V2.py

Thêm cờ --selftest: kiểm bản đã đóng gói có nạp đủ thư viện không rồi thoát.
CI chạy cờ này ngay sau khi build, để lỗi kiểu "Thiếu package supabase"
bị chặn tại chỗ thay vì tới tay người dùng.
"""
import sys


def _selftest():
    # Console Windows (cp1252/cp1258) không in được tiếng Việt: không ép UTF-8 thì
    # đúng lúc có lỗi thật, bước tự kiểm lại văng UnicodeEncodeError che mất lỗi.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    ok = True

    def report(name, good, err=''):
        nonlocal ok
        if good:
            print('OK    ' + name)
        else:
            print('FAIL  {} -> {}'.format(name, err or 'unknown'))
            ok = False

    import autopptx2.cloud as C
    print('RUNTIME ' + C.runtime_info())
    report('supabase', C.create_client is not None, C.SUPABASE_IMPORT_ERROR)
    report('keyring', C.keyring is not None, C.KEYRING_IMPORT_ERROR)

    # keyring 25 bỏ backend 'OS_X'. Nếu spec còn khai tên cũ, bản Mac đóng gói
    # sẽ thiếu backend mật khẩu -> "Nhớ mật khẩu" hỏng mà báo sai lỗi.
    if C.keyring is not None:
        try:
            _b = C.keyring.get_keyring()
            report('keyring backend ' + type(_b).__name__,
                   'fail' not in type(_b).__name__.lower())
        except Exception as e:
            report('keyring backend', False, '{}: {}'.format(type(e).__name__, e))

    for mod in ('tkinter', 'customtkinter', 'PIL.Image', 'pptx', 'openpyxl',
                'autopptx2.exporter', 'autopptx2.datasource'):
        try:
            __import__(mod)
            report(mod, True)
        except Exception as e:
            report(mod, False, '{}: {}'.format(type(e).__name__, e))

    try:
        import tkinter
        report('Tcl ' + tkinter.Tcl().eval('info patchlevel'), True)
    except Exception as e:
        report('Tcl', False, '{}: {}'.format(type(e).__name__, e))

    print('SELFTEST ' + ('PASS' if ok else 'FAIL'))
    return 0 if ok else 1


if __name__ == "__main__":
    if '--selftest' in sys.argv:
        sys.exit(_selftest())
    from autopptx2.app import main
    main()
