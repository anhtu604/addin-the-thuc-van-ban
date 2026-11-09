"""python va_ribbon.py <vao.dotm> <ra.dotm> : nhan Ribbon cho ban phat hanh."""
import sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data = zin.read(it.filename)
    if it.filename.startswith('customUI/') and it.filename.endswith('.xml'):
        t = data.decode('utf-8')
        old = 'supertip="Kiểm tra và cài phiên bản mới nhất từ GitHub. Thông tin đơn vị, khối đã chọn và bản quyền được giữ nguyên."'
        assert t.count(old) == 2
        t = t.replace(old, 'supertip="Kiểm tra và cài phiên bản mới nhất (người duy trì: Vương Đình Anh Tú). Thông tin đơn vị và khối đã chọn được giữ nguyên."')
        data = t.encode('utf-8')
    zout.writestr(it, data)
zout.close(); print('ribbon ok')
