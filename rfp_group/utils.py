import re

def expand_range(prefix, start, end):
    """주어진 접두사와 시작, 끝 번호로 범위를 확장하여 리스트를 반환합니다."""
    start_num = int(start)
    end_num = int(end)
    return [f"{prefix}{str(num).zfill(len(start))}" for num in range(start_num, end_num + 1)]

def parse_and_expand(input_str):
    """콤마로 구분된 문자열을 받아 확장된 리스트를 반환합니다."""
    input_str = input_str.replace(" ","")
    parts = input_str.split(',')
    expanded_list = []

    for part in parts:
        if '~' in part:
            match = re.match(r"([a-zA-Z]+-)(\d+)~(\d+)", part)
            if match:
                prefix, start, end = match.groups()
                expanded_list.extend(expand_range(prefix, start, end))
        else:
            match = re.match(r"([a-zA-Z]+-)?(\d+)", part)
            if match:
                prefix, number = match.groups()
                if prefix is None:
                    # 이전 접두사를 사용하여 번호를 확장합니다.
                    prev_prefix, prev_number = re.match(r"([a-zA-Z]+-)(\d+)", expanded_list[-1]).groups()
                    expanded_list.append(f"{prev_prefix}{number.zfill(len(prev_number))}")
                else:
                    expanded_list.append(f"{prefix}{number}")

    return expanded_list

def format_pages(pages):
    def page_key(page):
        match = re.match(r"(\D+)-(\d+)", page)
        return (match.group(1), int(match.group(2)))
    
    pages = sorted(pages, key=page_key)
    ranges = []
    range_start = pages[0]
    range_end = pages[0]
    prev_key = page_key(pages[0])

    for i in range(1, len(pages)):
        current_key = page_key(pages[i])
        if current_key[0] == prev_key[0] and current_key[1] == prev_key[1] + 1:
            range_end = pages[i]
        else:
            if range_start == range_end:
                ranges.append(prev_key[0] + '-' + str(page_key(range_start)[1]))
            else:
                ranges.append(f"{prev_key[0]}-{page_key(range_start)[1]}~{page_key(range_end)[1]}")
            range_start = pages[i]
            range_end = pages[i]
        prev_key = current_key
    
    if range_start == range_end:
        ranges.append(prev_key[0] + '-' + str(page_key(range_start)[1]))
    else:
        ranges.append(f"{prev_key[0]}-{page_key(range_start)[1]}~{page_key(range_end)[1]}")
    
    return ','.join(ranges)
