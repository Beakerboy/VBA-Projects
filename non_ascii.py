def find_non_ascii(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        for line_num, line in enumerate(f, 1):
            if not line.isascii():
                for char_num, char in enumerate(line, 1):
                    if not char.isascii():
                        print(f"Non-ASCII character '{char}' found at Line {line_num}, Position {char_num}")
