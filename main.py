import util
import os


def run():
    filespath = 'files'
    lists = os.listdir(filespath)
    print("发现文件：", len(lists), "个")
    for i, file in enumerate(lists):
        print(os.path.join(filespath, file))
        df = util.get_xlsx_to_dataframe(os.path.join(filespath, file))
        if (len(list(df.values)) > 0):
            fname, ext = os.path.splitext(file)
            util.dataframe_to_mssql(df, fname)
        print("当前进度：", i + 1,'/', len(lists))


if __name__ == '__main__':
    try:
        run()
    except Exception as e:
        print(e)
    finally:
        print("执行完毕")
