import qrcode
import excel
def get(src):

    for i in range(len(src)):

        data = src[i]
        # 实例化QRCode生成qr对象
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=3,
            border=4
        )
        # 传入数据
        qr.add_data(data)

        qr.make(fit=True)

        # 生成二维码
        img = qr.make_image()

        # 保存二维码
        img.save(r'dst' + '\\' + data + '.jpg')
        # 展示二维码
        # img.show()