while True:
    #V3匯出指定路徑資料夾內照片的EXIF及OS資訊(完整版)
    import os
    import exifread
    from PIL import Image
    from openpyxl import Workbook
    from openpyxl import load_workbook

    #創建excel
    wb = Workbook()
    ws = wb.active
    ws.title = 'img_info'
    title = ['檔名','檔名查驗','EXIF DateTimeOriginal','EXIF DateTimeDigitized','Image DateTime','相片尺寸(OS)','檔案類型','案件號']
    ws.append(title)

    path = input('請輸入路徑(複製照片-內容-位置路徑):')
    filelist = os.listdir(path)

    n = 0
    for i in filelist:
        filename = str(filelist[n])
        imgPath = path + os.sep + filelist[n]
        file = open(imgPath,'rb')
        tags = exifread.process_file(file) #process_file(f, stop_tag='UNDEF', details=True, strict=False, debug=False)
        #照片創建日期，將ImageDate轉換成日期(數字格式ex.1622465805.6013849)
        try:
            ExifDate = str(tags['EXIF DateTimeOriginal']).replace(':', '-', 2)
            ExifDTD = str(tags['EXIF DateTimeDigitized']).replace(':', '-', 2)
            ID = str(tags['Image DateTime']).replace(':', '-', 2)
        except:
            ExifDate=ExifDTD=ID=''
    #用try-except語法，當照片日期不存在時表格內空白
        img = Image.open(imgPath)
        w = img.width       #圖片的寬
        h = img.height      #圖片的高
        Image_size = str(w) + '*' + str(h)
        fortmat = img.format      #圖片格式
        filedata = ws.append([filename[:-4],'檔名查驗', ExifDate, ExifDTD, ID, Image_size, fortmat])
        n = n + 1
    wb.save('C:/Users/user/Desktop/檔案資料生成/xlsxname.xlsx')
    print("Finish")