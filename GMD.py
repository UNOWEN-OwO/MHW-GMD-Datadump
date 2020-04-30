#!/usr/bin/env python
# coding: utf-8

# By Jodo and Ice


import sys
import os
import re
import csv
import blowfish
import glob
from sys import argv
import pyexcel
from pyexcel.cookbook import merge_all_to_a_book


def find_file(file,folder):
    for root, dirs, files in os.walk(folder):
        if file in files:
            # print(root+'/'+file)
            return root+'/'+file
    print('File ' + file + ' not exits!')
    return None


def read_uint32(f):
    return ord(f.read(1)) + (ord(f.read(1)) << 8) + (ord(f.read(1)) << 16) + (ord(f.read(1)) << 24)


def read_uint16(f):
    return ord(f.read(1)) + (ord(f.read(1)) << 8)
    


def read_str(f):
    tmp = f.read(1)
    total = b''
    while(tmp is not None and ord(tmp) != 0):
        total += tmp
        tmp = f.read(1)
    return total.decode(encoding = 'utf-8').replace('\r\n', '')


def read_gmd(path):
    with open(path, 'rb') as f:
        id_list = []
        info_list = []
        empty_check = []
        f.read(20)
        count_id = read_uint32(f)
        count_info = read_uint32(f)
        length_id = read_uint32(f)
        length_info = read_uint32(f)
        length_file_name = read_uint32(f)
        f.read(length_file_name+1)
        for i in range(0, count_id):
            empty_check.append(read_uint32(f))
            f.read(28)
        # f.read(count_id << 5)
        f.read(2048)
        for i in range(0, count_info):
            if i not in empty_check:
                id_list.append('')
            else:
                id_list.append(read_str(f))
        for i in range(0, count_info):
            info_list.append(read_str(f))

        return id_list, info_list
    


def remove_style(lst):
    return [re.sub(r'<(?:\/STYL|STYL \w+)>', '',e) for e in lst]
    



def read_Askill(lang='eng', path='chunk/', out='ASkill_', filename='a_skill_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    gmd = list(zip(*read_gmd(filename)))
    
    fields = ['装备ID','装备代码','装备名称','强化装备名称','装备介绍',] if lang in ['chT', 'chS'] else ['Askill ID','Askill Code','Askill Name','Askill Upgrade Name','Askill Desc']
    
    rows = list(zip(
        [int(e[0].replace('ASkill', '').replace('_NAME', '')) - 1 for e in gmd if 'NAME' in e[0] and not 'UG' in e[0]],
        [e[0].replace('_NAME', '') for e in gmd if 'NAME' in e[0] and not 'UG' in e[0]],
        [e[1] for e in gmd if 'NAME' in e[0] and not 'UG' in e[0]],
        [e[1] for e in gmd if 'UG' in e[0]],
        [e[1] for e in gmd if e[0].endswith('EXPLAIN')]
    ))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Animal(lang='eng', path='chunk/', out='Animal_', filename=''):
    out = out + lang + '.tsv'
    
    animal_id = [0,1,1,2,2,3,3,4,5,5,5,5,6,6,7,7,7,8,8,9,9,10,10,10,11,11,12,12,12,12,13,13,14,15,15,16,17,18,18,18,19,19,20,20,21,22,23,23,23,24,24,25,26,27,28,29,30,31,31,31,31,31,32,32,32,32,32,33,33,34,34,35,35,36,37,37,37,37,38,38,39,39,40,40,41,41,42,43,43,43,44,50,50,51,52,53,54,55,56,57,57,58,59,59,59,59,59,59,59]
    animal_subid = [0,0,1,0,1,0,1,0,0,1,2,3,0,1,0,1,2,0,1,0,1,0,1,2,0,1,0,1,2,3,0,1,0,0,1,0,0,0,1,2,0,1,0,1,0,0,0,1,2,0,1,0,0,0,0,0,0,0,1,2,3,4,1,2,3,4,5,0,1,0,1,0,1,0,0,1,2,3,0,1,0,1,0,1,0,1,0,0,1,2,0,0,1,0,0,0,0,0,0,0,1,0,0,1,2,3,4,5,6]
    item_id = [-1,659,660,661,662,663,664,665,666,667,668,669,670,671,672,673,1792,674,675,676,945,677,678,1793,679,680,681,682,683,684,685,686,687,688,689,690,691,692,693,694,695,696,697,698,699,700,701,702,703,704,705,706,-1,-1,-1,-1,707,708,709,710,946,947,-1,-1,-1,-1,-1,711,712,713,714,715,716,-1,717,718,719,720,721,722,723,724,725,726,948,949,950,968,951,967,958,1455,1456,1458,1459,1460,1461,1462,1463,1464,1473,1465,1466,1467,1468,1469,1470,1471,1472]
    
    chT = ['Unavailable','繞行兔','嚮導兔','叢林始祖鳥','始祖鳥的使者','鈷閃蝶','幽閃蝶','淚之龜殼攀鱸','森林蜥蜴','蟻窩蜥蜴','黑暗蜥蜴','月光蜥蜴','霞霧女郎','斯卡班傑拉','憎恨鳥','桃源鄉','哨冰鸟','預兆蜻蜓','吉兆蜻蜓','鱗片打擊手','金色的鱗片打擊手','糞金龜','爆彈甲蟲','雪球虫','粉色異形魚','上等粉色異形魚','破裂龍魚','爆裂龍魚','上等破裂龍魚','上等爆裂龍魚','雅緻珊瑚鳥','伶俐珊瑚鳥','行燈魚','柔毛秧雞','粗毛秧雞','蛇麻古比魚','化石湯釜','麻痺瓦斯蛙','睡眠瓦斯蛙','爆炸瓦斯蛙','搖曳鰻','搖曳鰻女王','回復蜜蟲','大回復蜜蟲','飛天烏帽','搬運蟻','大獨角仙','金色獨角仙','虹色大獨角仙','皇帝蚱蜢','暴君蚱蜢','閃光翅蟲','羽虫','ハエ','ハチ','蛍','遠古鬼蝠魟','鐵鱟','軍團鱟','綠寶石鱟','金箔鱟','純金鱟','苦蟲','光蟲','不死蟲','雷光蟲','佳餚蟲','堅硬竹筴魚','大堅硬竹筴魚','貪吃鮪魚','貪吃鮪魚王','大王旗魚','帝王旗魚','楔虫','黃金魚','白金魚','大黃金魚','大白金魚','小金魚','小金魚王','刺身魚','大刺身魚','火藥金魚','大火藥金魚','銅饅頭蟹','金饅頭蟹','土龍','幼小仙人掌','仙人掌','幼小仙人花','孽鬼','溫泉銀猴','溫泉金猴','異形甲蟲','望月水母','胖衣企鵝','道標壁虎','小藍歌','神帶魚','玻璃異形魚','上等玻璃異形魚','月羽天母','毛吉','青苔毛吉','掘土毛吉','茸茸毛吉','刺針毛吉','火爆毛吉','白絲毛吉']    
    fields = ['生物ID','副ID','物品ID','名称'] if lang in ['chT', 'chS'] else ['EC ID','EC SubID','Item ID','EC Name']
    
    if lang not in ['chT', 'chS']:
        item = read_Item(lang)
        for i in item:
            if i[0] in item_id:
                chT[item_id.index(i[0])] = i[1]

    rows = list(zip(
        animal_id,
        animal_subid,
        item_id,
        chT
    ))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Armor(lang='eng', path='chunk/', out='Armor_', filename='armor_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    
    part_id = ['HEAD', 'BODY', 'ARM', 'WAIST', 'LEG', 'ACCE', 'DUMMY']
    part_name = ['头盔', '铠甲', '护手', '腰甲', '护腿', '护石', 'DUMMY'] if lang in ['chT', 'chS'] else ['Head','Body','Arm','Waist','Leg','Charm','DUMMY']
    part_path = ['helm', 'body', 'arm', 'wst', 'leg']
    gender = ['无', ' (男性限定)', ' (女性限定)', ''] if lang in ['chT', 'chS'] else ['None', ' (Male Only)', ' (Female Only)', '']
    
    fields = ['部位ID','部位名称','防具ID','防具名称','幻化ID','防御','稀有度','模型地址'] if lang in ['chT', 'chS'] else ['Equip Type ID', 'Equip Type','Equip ID', 'Equip Name', 'Transmog ID', 'Defense', 'Rare' ,'File Path']
    rows = []
    
    # print([e for e in gmd if part_id[0] in e[0]])
    
    with open(find_file('armor.am_dat', path),'rb') as armor:
        armor.read(6)
        cnt = read_uint32(armor)
        for i in range(cnt):
            idx = read_uint32(armor)
            order = armor.read(2)
            rank = armor.read(1)
            series_id = read_uint16(armor)
            is_layered = armor.read(1)
            armor_type = ord(armor.read(1))
            defense = read_uint16(armor)
            model_main_id = read_uint16(armor)
            model_sub_id = read_uint16(armor)
            armor.read(3)
            rare = ord(armor.read(1))
            armor.read(28)
            gender_id = read_uint32(armor)
            set_id = read_uint16(armor)
            armor.read(5)
            
            # print('%s%03d' % (part_id[armor_type],set_id))
            key = 'AM_%s%03d_NAME' % (part_id[armor_type],set_id)
            rows.append([armor_type, part_name[armor_type], set_id, info_list[id_list.index(key)] if key in id_list else 'Unavailable', series_id, defense, rare, ('pl%03d_%04d%s' % (model_main_id,model_sub_id,gender[gender_id])) if gender_id and armor_type < 5 else gender[0]])
            

    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_catSkill(lang='eng', path='chunk/', out='CatSkill_', filename='catSkill_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    
    fields = ['猫饭代码','猫饭名称','猫饭效果'] if lang in ['chT', 'chS'] else ['Meal Code', 'Meal Name', 'Meal Effect']
    
    rows = list(zip(id_list[0::2],info_list[0::2],info_list[1::2]))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Food(lang='eng', path='chunk/', out='Food_', filename='food_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    
    fields = ['食材代码','食材名称','食材介绍'] if lang in ['chT', 'chS'] else ['Food Code', 'Food Name', 'Food Desc']
    
    rows = list(zip(id_list[0::2],info_list[0::2],info_list[1::2]))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Gallery(lang='eng', path='chunk/', out='Gallery_', filename='cm_gallery_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    gmd = list(zip(id_list, info_list))

    fields = ['回放ID','回放名称','回放描述'] if lang in ['chT', 'chS'] else ['Gallery ID', 'Gallery Name','Gallery Desc']
    rows = [[e[0], e[1], gmd[id_list.index(e[0]+'_INFO')][1]] for e in gmd if 'EVC' in e[0] and not e[0].endswith('INFO')]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Item(lang='eng', path='chunk/', out='Item_', filename='item_'):
    out = out + lang + '.tsv'
    filename = filename + lang + '.gmd'
    
    item_id = []
    item_type = []
    item_rare = []
    with open(find_file('itemData.itm', path),'rb') as itemData:
        itemData.read(6)
        cnt = read_uint32(itemData)
        for i in range(cnt):
            item_id.append(read_uint32(itemData))
            itemData.read(1)
            item_type.append(read_uint32(itemData))
            item_rare.append(ord(itemData.read(1)))
            itemData.read(22)
    
    id_list = info_list = None
    for root, dirs, files in os.walk(path):
        if filename in files:
            p = root+'/'+filename
            id_list, info_list = read_gmd(p)
            if (len(id_list[0::2]) == len(item_rare)):
                break
    
    if not id_list or not info_list:
        print('File ' + filename + ' not exits!')
        return 
    
    type_list = ['物品','素材','换算道具','弹药或瓶','装饰品','房间装饰'] if lang in ['chT', 'chS'] else ['Item', 'Material', 'Account Item', 'Ammo or Coating', 'Jewel', 'Room Decoration']
    
    fields = ['物品ID','物品名称','物品类型','稀有度','物品介绍'] if lang in ['chT', 'chS'] else ['Item ID','Item Name','Item Type','Rare','Item Desc']
    rows = list(zip(
        item_id,
        info_list[0::2],
        [ type_list[e] for e in item_type],
        item_rare,
        info_list[1::2]))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Jewel(lang='eng', path='chunk/', out='Jewel_', filename='item_'):
    out = out + lang + '.tsv'
    item_data = read_Item(lang, path)
    if not item_data:
        return
    
    fields = ['装饰珠ID','物品ID','名称','孔位']  if lang in ['chT', 'chS'] else ['Deco ID', 'Item ID', 'Deco Name', 'Slot Size']
    rows = []
    
    with open(find_file('skillGemParam.sgpa', path),'rb') as jewelData:
        jewelData.read(6)
        cnt = read_uint32(jewelData)
        for i in range(cnt):
            item_id = read_uint32(jewelData)
            jewel_id = read_uint32(jewelData)
            size = read_uint32(jewelData)
            skill1 = read_uint32(jewelData)
            skill1_lv = read_uint32(jewelData)
            skill2 = read_uint32(jewelData)
            skill2_lv = read_uint32(jewelData)
            rows.append([jewel_id, item_id, item_data[item_id][1], size])

    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_lDelivery(lang='eng', path='chunk/', out='Delivery_', filename='l_delivery_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    _, info_list = read_gmd(filename)
    
    fields = ['交货委托ID','交货委托名称','交货委托回报'] if lang in ['chT', 'chS'] else ['Delivery ID', 'Delivery Name', 'Delivery Reward']
    rows = list(zip(range(len(info_list)),info_list[0::5],info_list[4::5]))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_lMission(lang='eng', path='chunk/', out='Mission_', filename='l_mission_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    _, info_list = read_gmd(filename)
    
    fields = ['奖金任务ID','奖金任务名称','奖金任务目标'] if lang in ['chT', 'chS'] else ['Mission ID', 'Mission Name','Mission Desc']
    rows = list(zip(range(len(info_list)),info_list[0::2],remove_style(info_list[1::2])))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Medal(lang='eng', path='chunk/', out='Achievement_', filename='cm_medal_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    gmd = list(zip(id_list, info_list))
    gmd = [[e[0], e[1] if e[0]+'_re' not in id_list else info_list[id_list.index(e[0]+'_re')]] for e in gmd if 'MEDAL' in e[0] and '_re' not in e[0] and 'set' not in e[0]]

    fields = ['成就ID','成就名称','成就要求'] if lang in ['chT', 'chS'] else ['Achieve ID', 'Achieve Name','Achieve Desc']
    rows = list(zip(
        [e[0].replace('MEDAL','').replace('_','') for e in gmd[0::3]],
        [e[1] for e in gmd[0::3]],
        [e[1] for e in gmd[1::3]],
    ))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Monster(lang='eng', path='chunk/', out='Monster_', filename='em_names_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    gmd = list(zip(*read_gmd(filename)))
    gmd = gmd[:114] + gmd[120:123] + [gmd[114]] + [gmd[123]] + [gmd[115]] + [gmd[124]] + [gmd[116]] + gmd[126:]
    
    fields = ['怪物ID','怪物代码','怪物名称'] if lang in ['chT', 'chS'] else ['Monster ID', 'Monster Code', 'Monster Name']
    rows = [[e[0], e[1][0], e[1][1]] for e in list(zip(list(range(len(gmd)//2)),gmd[0::2]))]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Music(lang='eng', path='chunk/', out='Music_', filename='music_skill_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    gmd = list(zip(*read_gmd(filename)))

    fields = ['音乐代码','音乐等级','音乐技能'] if lang in ['chT', 'chS'] else ['Music Code', 'Music Level','Music Name']
    rows = [[e[0], 0 if e[0].endswith('_S') else 1 if e[0].endswith('_W') else '-',e[1]] for e in gmd]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_OtArmor(lang='eng', path='chunk/', out='OtArmor_', filename='ot_armor_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    # gmd = [[e[0], e[1]] for e in list(zip(*read_gmd(filename))) if e[0].endswith('_NAME')]
    gmd = list(zip(*read_gmd(filename)))
    
    part_name = ['头盔', '铠甲'] if lang in ['chT', 'chS'] else ['Helmet', 'Armor']
    part_path = ['helm', 'body']
    
    fields = ['部位ID','部位名称','防具ID','防具名称','模型地址'] if lang in ['chT', 'chS'] else ['Equip Type ID', 'Equip Type','Equip ID', 'Equip Name', 'File Path']
    rows = []
    
    with open(find_file('otomoArmor.oam_dat', path),'rb') as otarmor:
        otarmor.read(6)
        cnt = read_uint32(otarmor)
        for i in range(cnt):
            otarmor.read(6)
            armor_type = ord(otarmor.read(1))
            otarmor.read(8)
            file_id = ord(otarmor.read(1))
            otarmor.read(20)
            armor_id = read_uint16(otarmor)
            name_id = read_uint16(otarmor)
            info_id = read_uint16(otarmor)
            
            rows.append([armor_type, part_name[armor_type], armor_id, gmd[name_id][1], 'otomo/equip/ot%03d/%s' % (file_id,part_path[armor_type])])

    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_OtWeapon(lang='eng', path='chunk/', out='OtWeapon_', filename='ot_weapon_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return

    gmd = list(zip(*read_gmd(filename)))
    
    fields = ['武器ID','武器名称','模型地址'] if lang in ['chT', 'chS'] else ['Weapon ID', 'Weapon Name','File Path']
    rows = []
    
    encrypt = find_file('otomoWeapon.owp_dat', path)
    with open(encrypt,'rb') as en, open(encrypt+'_d', 'wb') as de:
        cipher = blowfish.Cipher(b'FZoS8QLyOyeFmkdrz73P9Fh2N4NcTwy3QQPjc1YRII5KWovK6yFuU8SL',byte_order = 'little')
        de.write(b''.join(cipher.decrypt_ecb_cts(en.read())))
        
    with open(find_file('otomoWeapon.owp_dat_d', path),'rb') as otweapon:
        otweapon.read(6)
        cnt = read_uint32(otweapon)
        for i in range(cnt):
            index = read_uint32(otweapon)
            sid = read_uint32(otweapon)
            otweapon.read(15)
            file_id = read_uint32(otweapon)
            otweapon.read(5)
            weapon_id = read_uint16(otweapon)
            name_id = read_uint16(otweapon)
            info_id = read_uint16(otweapon)
            
            rows.append([weapon_id, gmd[name_id][1], 'otomo/wp/ot_we%03d' % file_id])
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


def read_Pugee(lang='eng', path='chunk/', out='Pugee_', filename='cm_pugee_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    gmd = list(zip(*read_gmd(filename)))

    fields = ['小猪服装ID','小猪服装名'] if lang in ['chT', 'chS'] else ['Pugee Cloth ID', 'Pugee Cloth Name']
    rows = [[e[0].replace('PUGEE_CLOTH_NAME_', ''),e[1]] for e in gmd if 'PUGEE_CLOTH_NAME_' in e[0]]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_Quests(lang='eng', path='chunk/', out='Quests_', filename='q'):
    out = out + lang + '.tsv'
    
    quest_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            #print(file, file[:len(filename)] == filename, )
            if file.startswith(filename) and file.endswith('_'+lang+'.gmd'):
                quest_list.append([file[len(filename):-len(lang)-5],root+'/'+file])
    
    fields = ['任务ID','任务名称','任务目标','任务失败条件'] if lang in ['chT', 'chS'] else ['Quest ID','Quest Name','Quest Target','Fail Condition']
    rows = [[quest[0]]+read_gmd(quest[1])[1][:3] for quest in quest_list]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_Skill(lang='eng', path='chunk/', out='Skill_', filename='skill_pt_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    
    fields = ['技能ID','技能代码','技能名称','技能介绍'] if lang in ['chT', 'chS'] else ['Skill ID','Skill Code','Skill Name', 'Skill Desc']
    rows = list(zip(
        range(len(id_list)//3),
        id_list[0::3],
        info_list[0::3],
        remove_style(info_list[2::3]),
    ))
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows
    


def read_Stage(lang='eng', path='chunk/', out='Stage_', filename='other_names_'):
    out = out + lang + '.tsv'
    filename = find_file(filename + lang + '.gmd', path)
    if not filename:
        return
    
    id_list, info_list = read_gmd(filename)
    gmd = list(zip(id_list, info_list))
        
    fields = ['场景ID','场景名称'] if lang in ['chT', 'chS'] else ['Stage ID','Stage Name']
    rows = [[e[0].replace('_NAME', ''), e[1]] for e in gmd if '_NAME' in e[0]]
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)
    
    return rows
    


def read_Weapon(lang='eng', path='chunk/', out='Weapon_', filename=''):
    out = out + lang + '.tsv'
    
    weapon_gmd = ['l_sword', 'sword', 'w_sword', 'tachi', 'hammer', 'whistle', 'lance', 'g_lance', 's_axe', 'c_axe', 'rod', 'bow', 'hbg', 'lbg']
    weapon_code = ['LSWD', 'SWD', 'WSWD', 'TACHI', 'HAMMER', 'WSL', 'LANCE', 'GLANCE', 'SAXE', 'CAXE', 'ROD', 'BOW', 'HBG', 'LBG']
    weapon_path = ['two','one','sou','swo','ham','hue','lan','gun','saxe','caxe','rod','bow','hbg','lbg']
    weapon_type = ['大剑','片手剑','双剑','太刀','大锤','狩猎笛','长枪','铳枪','斩斧','盾斧','操虫棍','弓','重弩炮','轻弩炮'] if lang in ['chT', 'chS'] else ['GreatSword','Sword&Shield','DualBlades','LongSword','Hammer','HuntingHorn','Lance','GunLance','SwitchAxe','ChargeBlade','InsectGlaive','Bow','HeavyBowGun','LightBowGun']
    dat_file = ['l_sword.wp_dat', 'sword.wp_dat', 'w_sword.wp_dat', 'tachi.wp_dat', 'hammer.wp_dat', 'whistle.wp_dat', 'lance.wp_dat', 'g_lance.wp_dat', 's_axe.wp_dat', 'c_axe.wp_dat', 'rod.wp_dat', 'bow.wp_dat_g', 'hbg.wp_dat_g', 'lbg.wp_dat_g']

    model_type = ['独立模型','组合模型'] if lang in ['chT', 'chS'] else ['Unique Model', 'Combined Model']
    sp_type = ['',' (凯罗武器)',' (冥赤武器)'] if lang in ['chT', 'chS'] else ['',' (KT Weapon)', ' (Safi Weapon)']
    sp_path = ['', 'bs_', '']
    
    all_id_list = []
    all_info_list = []
    for wg in weapon_gmd:
        filename = find_file(wg + '_' + lang + '.gmd', path)
        if not filename:
            return
        id_list, info_list = read_gmd(filename)
        all_id_list += id_list
        all_info_list += info_list
    
    fields = ['武器类型ID','武器类型','武器ID','武器名称','稀有度','模型类型','主模型地址','附件模型地址'] if lang in ['chT', 'chS'] else ['Weapon Type ID', 'Weapon Type', 'Weapon ID', 'Weapon Name', 'Rare','Model Type', 'Main Model Path', 'Sub Model Path']
    rows = []
    
    for wp_type_id in range(len(dat_file)):
        with open(find_file(dat_file[wp_type_id], path),'rb') as weapon:
            weapon.read(6)
            cnt = read_uint32(weapon)
            for i in range(cnt):
                idx = read_uint32(weapon)
                wp_id = read_uint16(weapon)
                unique_model = read_uint16(weapon)
                comb_model_main = read_uint16(weapon)
                comb_model_sub = read_uint16(weapon)
                weapon.read(8)
                rare = ord(weapon.read(1))
                if ('wp_dat_g' in dat_file[wp_type_id]):
                    weapon.read(40)
                else:
                    weapon.read(37)
                set_id = read_uint16(weapon)
                weapon.read(6)
            
                key = 'WP_%s_%03d_NAME' % (weapon_code[wp_type_id],set_id)
                rows.append([wp_type_id, 
                             weapon_type[wp_type_id], 
                             set_id, 
                             all_info_list[all_id_list.index(key)] if key in all_id_list else 'Unavailable', 
                             rare, 
                             '%s%s' % (model_type[0] if unique_model < 65535 else model_type[1], sp_type[wp_id//1000]),
                             'wp/%s/%s%s%03d' % (weapon_path[wp_type_id], sp_path[wp_id//1000], weapon_path[wp_type_id], unique_model if unique_model < 65535 else comb_model_main),
                             'wp/%s/parts/op_%s%03d' % (weapon_path[wp_type_id],weapon_path[wp_type_id],comb_model_sub) if comb_model_sub < 65535 else ''
                            ])
    
    with open(out, 'w', newline='',encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter='\t')
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

    return rows


# print('MHW 数据一键生成表格 V1.0 - 冰块\n')
print('MHW Data to Excel V1.0 - Ice\n')
path = 'chunk/'
lang = 'eng'
if len(argv) < 2:
    input('No folder input found, will use chunk/ folder in current directory,\nmake sure you include one with all files in chunk/common/\nPRESS ANY KEY\n')
else:
    path = argv[1]
print(['ara','chS','chT','eng','fre','ger','ita','jpn','kor','pol','ptB','rus','spa'])
lang = input('Type Language:\n')
if lang not in ['ara','chS','chT','DEV','eng','fre','ger','ita','jpn','kor','pol','ptB','rus','spa']:
    input('Language Code Invalid!\n')
    sys.exit()

read_list = [
    [ read_Askill, '特殊装备'],
    [ read_Animal, '环境生物'],
    [ read_Armor, '防具'],
    [ read_catSkill, '猫饭'],
    [ read_Food, '食材'],
    [ read_Gallery, '画廊'],
    [ read_Item, '物品'],
    [ read_Jewel, '装饰珠'],
    [ read_lDelivery, '交货任务'],
    [ read_lMission, '登录奖金'],
    [ read_Medal, '成就'],
    [ read_Monster, '怪物'],
    [ read_Music, '旋律'],
    [ read_OtArmor, '猫猫防具'],
    [ read_OtWeapon, '猫猫武器'],
    [ read_Pugee, '噗吱猪'],
    [ read_Quests, '任务'],
    [ read_Skill, '技能'],
    [ read_Stage, '场景'],
    [ read_Weapon, '武器'],
]

if lang in ['chT', 'chS']:
    for r, out in read_list:
        r(lang,path,out)
else:
    for r, out in read_list:
        r(lang,path)

merge_all_to_a_book([e for e in glob.glob('*.tsv') if lang in e], 'MHW '+ lang +'_GMD Data.xlsx')

print('\nMHW '+ lang +'_GMD Data.xlsx Generated!\n')


# re.sub(r'<(?:\/STYL|STYL \w+)>', '', 'LINE')



