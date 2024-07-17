import xlfr  # INFO: 240717 生成は、`maturin develop` (venv 等で該当環境に入ること。pip list で確認できる。)

if __name__ == '__main__':
    path = "./test.xlsm"
    data = xlfr.read_excel(path)

    for row in data:
        print(row)  # TODO: 240717 Rust -> Python の工程で、型変換できないか検討せよ。 (python 側でラップして、必要処理を加えれないか？)