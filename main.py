import openpyxl #エクセル操作
import shutil #フォルダ操作
from getpass import getpass #パスワード取得

#エクセルとの紐付け
kamoku = openpyxl.load_workbook('./kamoku.xlsx')
data = kamoku['data']

userdata = openpyxl.load_workbook('./userdata.xlsx')
userlist = userdata['userlist']

#基本機能
class func:

    #テキストの表示
    def txt(self, name):
        with open(f'txt/{name}.txt', 'r') as tmp:
            print(tmp.read(), end='')

    #excel/userdata.xlsx内のuserlistに保存
    def add_userlist(self, data, col, row):
        userlist[f'{col}{row}'] = data
        userdata.save('excel/userdata.xlsx')

func = func()

#ユーザーデータに関わる機能
class user:

    #ユーザーの新規登録
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

    #ログイン
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

    #ユーザーデータの削除
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

    #コースの入力
    def corse(self, input1, inputsave):
        while(True):
            print('コースを入力してください')
            for i in range(1,5):
                print(input1[f'd{i}'].value + ':' + input1[f'e{i}'].value)
            input1['b1'].value = int(input())
            for i in range(0,4):
                if input1['b1'].value == i:
                    inputsave()
                    print('COMPLETED: 正常に登録されました')
                    return 0
            input1['b1'].value = None
            print('ERROR：もう一度最初から入力してください')

    #コースの表示
    def corse_show(self, input1):
        for i in range(1, 5):
            if input1[f'b1'].value == input1[f'd{i}'].value:
                return input1[f'e{i}'].value

    #学年の入力
    def grade(self, input1, inputsave):
        while(True):
            print('学年を入力してください')
            input1['b2'].value = int(input())
            if input1['b2'].value in {1, 2, 3, 4}:
                inputsave()
                print('COMPLETED: 正常に登録されました')
                return 0
            input1['b1'].value = None
            print('ERROR：もう一度最初から入力してください')

    #学年の表示
    def grade_show(self, input1):
        if int(input1['b2'].value) in {1,2,3,4}:
            return int(input1['b2'].value)
        return 0
    
    #教職の入力
    def tp(self, input1, inputsave):
        print('受けているものには1、そうでないものには0を入力してください')
        for i in range(3, 8):
            print(input1[f'a{i}'].value + ':', end="")
            while(True):
                input1[f'b{i}'].value = int(input())
                if input1[f'b{i}'].value in {0, 1}:
                    break
                print('ERROR：もう一度入力し直してください')
        print('COMPLETED: 正常に登録されました')
        return 0
                

user = user()

class calc:
    def gpa(self, input3, data):
        sum1 = 0
        sum2 = 0
        for i in range(2,166):
            sum1 = sum1 + data[f'c{i}'].value * (input3[f'd{i}'].value*4 + input3[f'e{i}'].value*3 + input3[f'f{i}'].value*2 + input3[f'g{i}'].value*1)
            sum2 = sum2 + data[f'c{i}'].value * (input3[f'd{i}'].value + input3[f'e{i}'].value + input3[f'f{i}'].value + input3[f'g{i}'].value + input3[f'h{i}'].value)
        return float(sum1/sum2)

    def gpaall(self):
        return 0

    def labo(self):
        return 0

    def graduate(self):
        return 0

calc = calc()

def reset():
    pass

def main():
    while(True):
        func.txt('start')
        c1 = int(input())
        if c1 in {1, 2, 3, 4}:
            #1: アカウント新規作成
            if c1 == 1:
                while(True):
                    if user.new() == 1:
                        break
            #2: ログイン
            elif c1 == 2:
                logindata = user.login()
                if logindata != 0:
                    #ユーザーデータの読み込み
                    input_ = openpyxl.load_workbook(f'./data/{logindata[1]}/input.xlsx')
                    input1 = input_['input1']
                    input2 = input_['input2']
                    input3 = input_['input3']
                    input4 = input_['input4']
                    def inputsave():    #保存用関数（引数として使う）
                        input_.save(f'./data/{logindata[1]}/input.xlsx')
                    print(f'ようこそ {logindata[1]} さん')
                    #ここにログイン後の操作を記述予定
                    while(True):
                        func.txt('after_login')
                        c2 = int(input())
                        if c2 in {1, 2, 3, 4, 5, 6}:
                            #1: ユーザー情報・成績確認
                            if c2 == 1:
                                print('ユーザーネーム: '+f'{logindata[1]}')
                                print('コース: ')
                                pass
                            #2: アカウント情報変更
                            elif c2 == 2:
                                pass
                            #3: 理学部科目入力
                            elif c2 == 3:
                                pass
                            #4: 教職科目入力
                            elif c2 == 4:
                                pass
                            #5: その他の科目入力
                            elif c2 == 5:
                                pass
                            #
                            elif c2 == 6:
                                print('ログアウトしました。')
                                break
                        else:
                            print('ERROR: もう一度入力し直してください')
                    #print(calc.gpa(input3, data))
                    #user.corse(input1, inputsave)
                    break
            #3: アカウント削除
            elif c1 == 3:
                while('True'):
                    if user.delete() >= 1:
                        break
            #4: 終了する
            elif c1 == 4:
                print('**終了**')
                break
        else:
            print('ERROR: もう一度入力し直してください')

if __name__ == "__main__":
    main()