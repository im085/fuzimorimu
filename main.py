import openpyxl
import shutil
from getpass import getpass
import codecs

#エクセルとの紐付け
kamoku = openpyxl.load_workbook('excel/kamoku.xlsx')
data = kamoku['data']

userdata = openpyxl.load_workbook('excel/userdata.xlsx')
userlist = userdata['userlist']

input_ = openpyxl.load_workbook('copy/input.xlsx')
input1 = input_['input1']
input2 = input_['input2']
input3 = input_['input3']

class func:
    def txt(self, name):
        with open(f'txt/{name}.txt', 'r') as tmp:
            print(tmp.read(), end='')

    def add_userlist(self, data, col, row):
        userlist[f'{col}{row}'] = data
        userdata.save('excel/userdata.xlsx')
func = func()

class user:
    def new(self):
        i = 2
        print('希望するユーザーネームを入力してください: ', end="")
        username = input()
        password = getpass()
        #ユーザーネームの重複がないか確認、重複の場合0を返す
        while(userlist[f'b{i}'].value != None):
            if username == userlist[f'b{i}'].value:
                print('ERROR: このユーザーネームは使われています')
                return 0
            i = i + 1
        #重複がなければ保存とアカウント用のディレクトリを作成し1を返す
        print('COMPLETED: ' + username + 'さん、ユーザー登録が完了しました')
        func.add_userlist(i-1, 'a', i)
        func.add_userlist(username, 'b', i)
        func.add_userlist(password, 'c', i)
        shutil.copytree('./copy', './data/'+f'{username}')
        return 1

    def login(self):
        i = 2
        print('ユーザーネームを入力してください: ', end="")
        username = input()
        password = getpass()
        while(userlist[f'b{i}'].value != None):
            if username == userlist[f'b{i}'].value and password == userlist[f'c{i}'].value:
                return [userlist[f'a{i}'].value, username, password]
            i = i + 1
        print('ERROR: ユーザーネームかパスワードが間違っています')
        return 0

    def delete(self):
        i = 2
        print('削除したいアカウントのユーザーネームを入力してください: ', end="")
        username = input()
        password1 = getpass()
        while(userlist[f'b{i}'].value != None):
            if username == userlist[f'b{i}'].value and password1 == userlist[f'c{i}'].value:
                print('本当に削除しますか？削除する場合、【y】と入力してください: ', end="")
                c = input()
                if c == 'y':
                    print('確認のためもう一度パスワードを入力してください')
                    password2 = getpass()
                    if password1 == password2:
                        print('アカウントを削除します')
                        #アカウント削除の処理をする
                        userlist.delete_rows(i)
                        shutil.rmtree('./data/'+f'{username}')
                        userdata.save('excel/userdata.xlsx')
                        return 1
                    else:
                        print('パスワードが間違っています')
                        return 0
                else:
                    print('スタートメニューに戻ります')
                    return 2
            i = i + 1
        print('一致するデータがありませんでした、スタートメニューに戻ります')
        return 2

    def main(self):
        return 0
user = user()

def corse():
    #コースの入力
    while(True):
        print('コースを入力してください')
        for i in range(1,5):
            print(input1[f'd{i}'].value,input1[f'e{i}'].value)
        input1['b1'].value = int(input())
        for i in range(0,4):
            if input1['b1'].value == i:
                return i
        input1['b1'].value = None
        print('ERROR：もう一度入力してください')
        print('')

class Calc:
    def __init__(self, name):
        self.name = self
    
    def gpa(self):
        return 0

    def gpaall(self):
        return 0

    def labo(self):
        return 0

    def graduate(self):
        return 0

def reset():
    pass

def main():
    while(True):
        func.txt('start')
        c = int(input())
        if c == 1:
            while('True'):
                if user.new() == 1:
                    break
        elif c == 2:
            logindata = user.login()
            if logindata != 0:
                break
        elif c == 3:
            while('True'):
                if user.delete() >= 1:
                    break
        elif c == 4:
            print('**終了**')
            break
        else:
            print('ERROR: もう一度入力し直してください')

if __name__ == "__main__":
    main()