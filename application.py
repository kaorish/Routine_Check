from routine_check_whole import main as fun1
from routine_check import main as fun2
from notification import main as fun3


def print_main():
    print('=' * 51)
    print("{:-^32}".format('欢迎来到江西理工大学信息工程学院查寝表格制作程序'))
    print("{:<48}".format('0.退出'))
    print("{:<43}".format('1.生成常规检查（whole）'))
    print("{:<37}".format('2.生成常规检查（需先完成第一步）'))
    print("{:<38}".format('3.生成周通报（需先完成第二步）'))
    print("{:<33}".format('4.生成每月大功率及安全隐患汇总表（未完成）'))
    print('-' * 51)


while True:
    print_main()
    a = input("请输入：")
    if a == "0":
        print('exit')
        break
    elif a == "1":
        fun1()
    elif a == "2":
        fun2()
    elif a == "3":
        fun3()
    elif a == "4":
        print('此功能未完成，待续...')
        continue
