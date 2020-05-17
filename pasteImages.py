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
    'img_margin_bottom': 4,
    'indent_size': 3
}
# 小項目に使える行数
config['seg_row_num'] = (config['page_row_num'] - config['page_padding']) // config['page_seg_num']
# 画像の縦の行数
config['img_row_num'] = config['seg_row_num'] - config['img_margin_bottom'] - 2
# 画像の縦の長さ（1行 == 20px）
config['img_hight'] = config['img_row_num'] * 20

def insert_image(img_path, insert_row_no):
    im = Image.open(img_path)
    img_hight = im.size[0]
    img_scale = config['img_hight'] / img_hight

    # 画像は横幅が少し長くなるため調整するマジックナンバーを定義
    adjust_scale = 0.82

    img_scale_tupple = {
        'x_scale': img_scale * adjust_scale,
        'y_scale': img_scale
    }

    worksheet.insert_image('D{0}'.format(insert_row_no), img_path, img_scale_tupple)

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
    primary_title_flag = 0
    for text in data.text:
        img_path = 'image/{0}'.format(img_names[count])

        page_no = count // config['page_seg_num']
        seg_no = count % config['page_seg_num']

        if primary_title_flag == 0 and seg_no == 0:
            page_margin_top = 1
        else:
            primary_title_flag = 0

        # セグメントの開始行を計算
        insert_row_no = page_no*config['page_row_num'] + \
            seg_no*config['seg_row_num'] + \
            page_margin_top + 1

        print(text)
        print(insert_row_no)
        
        # データの中身が大項目のタイトルデータなら別の処理
        if 'primary-title' in text:
            worksheet.merge_range('B{0}:K{0}'.format(insert_row_no), \
                text['primary-title'], primary_title_format)
            worksheet.set_row(insert_row_no-1, 25)
            page_margin_top += 2
            primary_title_flag = 1
            continue

        worksheet.merge_range('C{0}:K{0}'.format(insert_row_no), \
            text['title'], title_format)
        worksheet.set_row(insert_row_no-1, 20)
        insert_row_no += 2

        # 画像の挿入
        insert_image(img_path, insert_row_no)

        count += 1
    
    workbook.close()

