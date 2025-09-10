import re


def step1_split(address):
    # 定義正則表達式模式
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(?:、|/|與|及)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    # 使用正則表達式進行匹配
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(4)
        note = match.group(5) if match.group(5) else ''
        print(match, "1")
        return city, district, road1, road2, note
    else:
        return None


def step2_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(?:、|/|與|及)(\D{3,10})(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(4)
        note = match.group(5) if match.group(5) else ''
        print(match, "2")
        return city, district, road1, road2, note
    else:
        return None


def step3_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{3,10})(?:、|/|與|及)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(4)
        note = match.group(5) if match.group(5) else ''
        print(match, "3")
        return city, district, road1, road2, note
    else:
        return None


def step4_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(4)
        note = match.group(5) if match.group(5) else ''
        print(match, "4")
        return city, district, road1, road2, note
    else:
        return None


def step5_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{1,9}線)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(4)
        note = ''
        print(match, "5")
        return city, district, road1, road2, note
    else:
        return None


def step6_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(國道([1-9]|10)號)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = match.group(5)
        note = ''
        print(match, "6")
        return city, district, road1, road2, note
    else:
        return None


def step7_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = ''
        note = match.group(4)
        print(match, "7")
        return city, district, road1, road2, note
    else:
        return None


def step8_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(\D{3,10})\((.*)\)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = match.group(3)
        road2 = ''
        note = match.group(4)
        print(match, "8")
        return city, district, road1, road2, note
    else:
        return None


def step9_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{2}市|\D{1,3}(?:區|鄉|鎮))(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(2)
        road1 = ''
        road2 = ''
        note = match.group(3)
        print(match, "9")
        return city, district, road1, road2, note
    else:
        return None


def step10_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(?:、|/|與|及)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = match.group(3)
        note = match.group(4) if match.group(4) else ''
        print(match, "10")
        return city, district, road1, road2, note
    else:
        return None


def step11_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(?:、|/|與|及)(\D{3,10})(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = match.group(3)
        note = match.group(4) if match.group(4) else ''
        print(match, "11")
        return city, district, road1, road2, note
    else:
        return None


def step12_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{3,10})(?:、|/|與|及)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = match.group(3)
        note = match.group(4) if match.group(4) else ''
        print(match, "12")
        return city, district, road1, road2, note
    else:
        return None


def step13_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = match.group(3)
        note = match.group(4) if match.group(4) else ''
        print(match, "13")
        return city, district, road1, road2, note
    else:
        return None


def step14_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{1,9}線)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = match.group(3)
        note = ''
        print(match, "14")
        return city, district, road1, road2, note
    else:
        return None


def step15_split(address):
    pattern = r"(\D{2}(?:市|縣))(\S{2,5}(?:路|街|道)(?:[\d一二三四五六七八九十]+段)?)(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = ''
        note = match.group(3)
        print(match, "15")
        return city, district, road1, road2, note
    else:
        return None


def step16_split(address):
    pattern = r"(\D{2}(?:市|縣))(\D{3,10})\((.*)\)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = match.group(2)
        road2 = ''
        note = match.group(3)
        print(match, "16")
        return city, district, road1, road2, note
    else:
        return None


def step17_split(address):
    pattern = r"(\D{2}(?:市|縣))(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = ''
        road1 = ''
        road2 = ''
        note = match.group(2)
        print(match, "17")
        return city, district, road1, road2, note
    else:
        return None


def step18_split(address):
    pattern = r"(\D{3}(?:號))(\D{3}(?:號))(.*)"
    match = re.match(pattern, address)
    if match:
        city = match.group(1)
        district = match.group(1)
        road1 = match.group(2)
        road2 = match.group(3)
        note = ''
        print(match, "18")
        return city, district, road1, road2, note
    else:
        return None
