import xlfr  # INFO: 240717 生成は、`maturin develop` (venv 等で該当環境に入ること。pip list で確認できる。)

if __name__ == '__main__':
    path = "./test.xlsm"
    data = xlfr.read_excel(path)  # TODO: 240717 型ヒントが出るようにせよ。(polars とかはさすがに出ると思うので、参考にせよ。)

    for row in data:
        print(row)  # TODO: 240717 Rust -> Python の工程で、型変換できないか検討せよ。