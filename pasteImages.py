import os
import xlsxwriter
from PIL import Image

# 入力データ
import data

workbook = xlsxwriter.Workbook('out.xlsx')
worksheet = workbook.add_worksheet()

config = {
    'page_row_num': 57,
    'page_padding': 4,
    'page_seg_num': 2,
    'indent_size': 3
}
config['seg_row_num'] = (config['page_row_num'] - config['page_padding']) // config['page_seg_num']

def insert_image(img_path, insert_row_no):
    im = Image.open(img_path)
    # print(im.size)

    # 画像は横幅が少し長くなるため調整するマジックナンバーを定義
    width_additional_scale = 0.82

    img_scale = {
        'x_scale': 1 * width_additional_scale,
        'y_scale': 1
    }

    worksheet.insert_image('D{0}'.format(insert_row_no), img_path, img_scale)

if __name__ == "__main__":
    worksheet.set_column('A:C', config['indent_size'])
    img_names = os.listdir('image')

    # 大項目用フォーマット
    primary_title_format = workbook.add_format({
        'bold': 1,
        # 'border': 1,
        # 'align': 'center',
        'valign': 'vcenter',
        'font_color': 'white',
        'fg_color': 'black'})
    
    # 小項目用フォーマット
    title_format = workbook.add_format({
        'bold': 1,
        'valign': 'vcenter',
        'fg_color': '#BFBFBF'})

    count = 0
    page_margin_top = 1
    for text in data.text:
        img_path = 'image/{0}'.format(img_names[count])

        # セグメントの開始行を計算
        page_no = count // config['page_seg_num']
        seg_no = count % config['page_seg_num']
        insert_row_no = page_no*config['page_row_num'] + \
            seg_no*config['seg_row_num'] + \
            page_margin_top + 1

        
        # データの中身が大項目のタイトルデータなら別の処理
        if 'primary-title' in text:
            worksheet.merge_range('B{0}:K{0}'.format(insert_row_no), \
                text['primary-title'], primary_title_format)
            worksheet.set_row(insert_row_no-1, 25)
            page_margin_top += 2
            continue

        worksheet.merge_range('C{0}:K{0}'.format(insert_row_no), \
            text['title'], title_format)
        worksheet.set_row(insert_row_no-1, 20)
        insert_row_no += 2

        # 画像の挿入
        insert_image(img_path, insert_row_no)

        count += 1
        page_margin_top = 1
    
    workbook.close()

