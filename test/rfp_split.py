import re

def expand_range(prefix, start, end):
    start_num = int(start)
    end_num = int(end)
    return [f"{prefix}{str(num).zfill(len(start))}" for num in range(start_num, end_num + 1)]

def parse_and_expand(input_str):
    parts = input_str.split(',')
    expanded_list = []

    for part in parts:
        if '~' in part:
            prefix, range_part = re.match(r"([a-zA-Z]+-)(\d+~\d+)", part).groups()
            start, end = range_part.split('~')
            expanded_list.extend(expand_range(prefix, start, end))
        else:
            match = re.match(r"([a-zA-Z]+-)?(\d+)", part)
            if match:
                prefix, number = match.groups()
                if prefix is None:
                    # Handle case like "001,002" (continuation of previous prefix)
                    prev_prefix, prev_number = re.match(r"([a-zA-Z]+-)(\d+)", expanded_list[-1]).groups()
                    expanded_list.append(f"{prev_prefix}{number.zfill(len(prev_number))}")
                else:
                    expanded_list.append(f"{prefix}{number}")

    return expanded_list

# 예시 입력
examples = [
    "aaa-001",
    "aaa-001~004",
    "aaa-001,003",
    "aaa-001,bbb-001~003",
    "bbb-001~003,ccc-001,004",
    "DQR-005,DAR-001,009,013"
]

# 각 예시를 처리 및 출력
for idx, example in enumerate(examples, start=1):
    result = parse_and_expand(example)
    print(f"{idx}) {example}")
    print(f"답: {','.join(result)}")
    print()
