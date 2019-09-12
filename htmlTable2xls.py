# -*-coding:utf8-*-
# Author: KertinH
# Date: 2019/09/12


import xlwt
import os


def htmlTable2xls(htmlTable_list, save_file_path, file_name):
    if len(htmlTable_list) > 0:

        try:
            os.makedirs(save_file_path)
        except:
            pass

        count = 1
        workbook = xlwt.Workbook()
        for table in htmlTable_list:
            worksheet = workbook.add_sheet('test_sheet_{}'.format(count))
            count += 1
            # 记录所有cell的左上、右上、左下、右下点坐标  例：cell_lup = [row, col]
            cell_position = {}
            tr_num = len(table.xpath('.//tr'))
            # print(tr_num)
            true_tr = [100000000]
            # 解析tr标签
            for i in range(1, tr_num + 1):
                col = 0
                cell_lup = [ ]
                cell_rup = [ ]
                cell_ldown = [ ]
                cell_rdown = [ ]
                if table.xpath('.//tr')[i - 1].xpath('.//td'):
                    true_tr.append(i)
                    min_tr = min(true_tr)
                    td_num = len(table.xpath('.//tr')[i - 1].xpath('.//td'))
                    # print(td_num)
                    # 解析td标签
                    for j in range(1, td_num + 1):
                        row_now = table.xpath('.//tr')[i - 1].xpath('.//td')[j - 1].xpath('./@rowspan')
                        col_now = table.xpath('.//tr')[i - 1].xpath('.//td')[j - 1].xpath('./@colspan')
                        content = ''.join(
                            table.xpath('.//tr')[i - 1].xpath('.//td')[j - 1].xpath('.//text()')
                        ).replace('\n', '').replace('\u3000', '').replace('\xa0', '').replace(r'\&nbsp;', '')

                        # 获取跨行跨列数（rowspan、colspan）
                        if not row_now:
                            row_span = 1
                        else:
                            if type(row_now) == list:
                                row_now = int(''.join(row_now))
                            row_span = row_now
                        if not col_now:
                            col_span = 1
                        else:
                            if type(col_now) == list:
                                col_now = int(''.join(col_now))
                            col_span = col_now
                        col += col_span
                        # 处理表格第一行，根据跨行跨列数计算出cell坐标
                        if i == min_tr and j == 1:
                            cell_lup = [0, 0]
                            cell_rup = [0, col]
                            cell_ldown = [row_span, 0]
                            cell_rdown = [row_span, col]
                        elif i == min_tr and j != 1:
                            cell_lup = [0, col - col_span]
                            cell_rup = [0, col]
                            cell_ldown = [row_span, col - col_span]
                            cell_rdown = [row_span, col]
                        # 处理除第一行外的其它行，根据之前存在的cell坐标计算出自身坐标
                        if i != 1:
                            position_li = []
                            # 生成包含目前所有已知cell的 键、左下坐标 组成的列表
                            for position in cell_position.keys():
                                position_li.append([position, cell_position[position]['ld'],
                                                    cell_position[position]['rd']])
                            position_li = sorted(position_li, key=lambda x: (x[1][0], x[1][1]))
                            # 获取cell_position的所有键，cell_position为包含目前所有cell的属性集合（四角坐标、内容）
                            positions = [a for a in cell_position.keys()]
                            # 遍历cell_position的键，判断当前cell左上坐标位置，并以此结合跨行跨列值计算出当前cell的四角坐标
                            for position in positions:
                                cell_lup = position_li[0][1]
                                cell_rup = [position_li[0][1][0], position_li[0][1][1] + col_span]
                                cell_ldown = [position_li[0][1][0] + row_span, position_li[0][1][1]]
                                cell_rdown = [position_li[0][1][0] + row_span, position_li[0][1][1] + col_span]
                                if position == '{}_{}'.format(i, j - 1):
                                    # 根据当前cell的同行的前一个cell的右下角坐标获取前一cell的长度，判断前一cell长度是否超出其跟随的前一行cell长度
                                    # 若是，则移除当前前代cell坐标，获取新的前代cell坐标，以此计算当前cell四角坐标
                                    # 否则，取前一cell的右上角坐标作为当前cell的左上角坐标，计算当前cell四角坐标
                                    position_li_len = len(position_li)
                                    for num in range(1, position_li_len + 1):
                                        if cell_position['{}_{}'.format(i, j - 1)]['rd'][1] >= position_li[0][2][1] \
                                                and cell_position['{}_{}'.format(i, j - 1)]['rd'][0] > \
                                                position_li[0][2][0] and i != int(position_li[0][0].split('_')[0]):
                                            del cell_position[position_li[0][0]]
                                            del position_li[0]
                                            cell_lup = position_li[0][1]
                                            cell_rup = [position_li[0][1][0], position_li[0][1][1] + col_span]
                                            cell_ldown = [position_li[0][1][0] + row_span, position_li[0][1][1]]
                                            cell_rdown = [
                                                position_li[0][1][0] + row_span, position_li[0][1][1] + col_span]
                                        else:
                                            if cell_position['{}_{}'.format(i, j - 1)]['rd'][1] < position_li[0][1][1]:
                                                break
                                            cell_lup = cell_position['{}_{}'.format(i, j - 1)]['ru']
                                            cell_rup = [cell_position['{}_{}'.format(i, j - 1)]['ru'][0],
                                                        cell_position['{}_{}'.format(i, j - 1)]['ru'][1] + col_span]
                                            cell_ldown = [cell_position['{}_{}'.format(i, j - 1)]['ru'][0] + row_span,
                                                          cell_position['{}_{}'.format(i, j - 1)]['ru'][1]]
                                            cell_rdown = [cell_position['{}_{}'.format(i, j - 1)]['ru'][0] + row_span,
                                                          cell_position['{}_{}'.format(i, j - 1)]['ru'][1] + col_span]
                                            break
                                    # 若当前cell为本行最后一个cell，且其长度已超出或等于前代cell长度，移除前代cell
                                    if j == td_num:
                                        for index in position_li:
                                            if cell_rdown[1] >= index[2][1] and cell_rdown[0] > index[2][0] and \
                                                    i != int(index[0].split('_')[0]):
                                                del cell_position[index[0]]
                                            # 若前一行cell绝对宽度 > 本行最后一个cell绝对宽度的，移除
                                            if cell_rdown[1] < index[2][1] and i - 1 == int(index[0].split('_')[0]):
                                                del cell_position[index[0]]
                                    break
                                # 若td_num = 1时，当前cell宽大于或等于所有前代cell宽，则清空当前cell_position
                                # 否则遍历cell_position,移除其中绝对宽度小于当前cell的前代cell
                                rows = [0]
                                cols = [0]
                                for data in position_li:
                                    rows.append(data[1][0])
                                    cols.append(data[1][1])
                                if 1 == td_num:
                                    if cell_rdown[0] > max(rows) and col_span >= max(cols):
                                        cell_position = {}
                                        break
                                    else:
                                        position_li_len = len(position_li)
                                        for num in range(1, position_li_len + 1):
                                            if cell_rdown[1] >= position_li[0][2][1]:
                                                # print(cell_position, '    161')
                                                del cell_position[position_li[0][0]]
                                                del position_li[0]
                                        break
                        # 记录cell的编号、四角坐标、内容
                        cell_position['{}_{}'.format(i, j)] = {
                            'lu': cell_lup, 'ru': cell_rup, 'ld': cell_ldown, 'rd': cell_rdown, 'content': content}
                        worksheet.write_merge(cell_lup[0], cell_rdown[0] - 1,
                                              cell_lup[1], cell_rdown[1] - 1, content)
                        pass
            workbook.save(save_file_path + '\\{}.xls'format(file_name))
