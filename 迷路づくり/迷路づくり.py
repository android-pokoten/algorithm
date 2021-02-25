import xlwings as xw
import random

# パラメータ
# 最大範囲 (奇数のみ)
maxrange = 25
# 壁の色(白)
color_wall = (60, 180, 0)
# 道の色(薄い灰色)
color_road = (200, 200, 255)
# スタート色
color_start = (255, 0, 0)
# ゴール色
color_finish = (0, 0, 255)
# open色
color_open = (100, 100, 255)
# close色
color_close = (100, 100, 100)
# 最短ルート色
color_route = (0, 255, 0)
# 列幅
col_width = 3
# 行高さ
row_height = 18.75

# 変数定義
# 上へ進める
can_up = 1
# 左へ進める
can_left = 2
# 下へ進める
can_down = 3
# 右へ進める
can_right = 4
# どこへも進めない
cannot_turn = 9

# 作成完了文字列
finish_strings = 'finish'
# 文字サイズ初期値
default_fontsize = 11

# open セルオブジェクト
opens = []

# 行先判定
# 　上下左右で行けるところを返す
#　　行けるところ = 枠外にならないか、すでに道が引かれていない
def check(row, col):
    rets = []
    # ↑
    if row > 2:
        if xw.Range((row - 2, col)).color == color_wall:
            rets.append(can_up)

    # ←
    if col > 2:
        if xw.Range((row, col - 2)).color == color_wall:
            rets.append(can_left)

    # ↓
    if maxrange > row + 2:
        if xw.Range((row + 2, col)).color == color_wall:
            rets.append(can_down)

    # →
    if maxrange > col + 2:
        if xw.Range((row, col + 2)).color == color_wall:
            rets.append(can_right)

    if len(rets) == 0:
        # どこにも行けない場合
        return cannot_turn
    # 行ける場合はランダムで方向を返す
    return random.choice(rets)

def make_maize():
    # エリアを初期化
    xw.Range((1, 1), (maxrange, maxrange)).value = ""
    xw.Range((1, 1), (maxrange, maxrange)).color = color_wall
    xw.Range((1, 1), (maxrange, maxrange)).row_height = row_height
    xw.Range((1, 1), (maxrange, maxrange)).column_width = col_width
    xw.Range((1, 1), (maxrange, maxrange)).api.Font.Size = default_fontsize
    # 初期地点をランダムに選ぶ(偶数マスである必要あり)
    row = random.randint(1, (maxrange // 2)) * 2
    col = random.randint(1, (maxrange // 2)) * 2
    # 初期地点をスタックにセットし、色を塗っておく
    stack = [[row, col]]
    xw.Range((row, col)).color = color_road

    # スタックが空=スタート地点まで戻ってきた、となるまでループ
    while stack:
        row, col = stack.pop()
        ret = 0

        while ret != cannot_turn:
            ret = check(row, col)

            if ret == can_up:
                rdiff = -2
                cdiff = 0

            if ret == can_left:
                rdiff = 0
                cdiff = -2

            if ret == can_down:
                rdiff = 2
                cdiff = 0

            if ret == can_right:
                rdiff = 0
                cdiff = 2
            
            # どこへも行けない場合は色塗りや現在位置のスタックセットをしない
            if ret != cannot_turn:
                xw.Range((row, col), (row + rdiff, col + cdiff)).color = color_road
                # 現在位置をスタックに追加した上で更新
                stack.append([row, col])
                row += rdiff
                col += cdiff
    
    xw.Range((1, 1)).value = finish_strings
    xw.Range((1, 1)).select()

def get_opens():
    ret = []
    # color_open のセルを探す
    rng = xw.Range((2,2), (maxrange - 1, maxrange - 1))
    for i in rng:
        if i.color == color_open:
            ret.append(i)

    # もし該当するセルが0の場合、スタート直後とみなしてcolor_start のセルを返す
    if len(ret) == 0:
        ret.append(xw.Range((2, 2)))

    return ret

def heuristic_cost(row, col):
    # 予想コストを算出
    # とりあえずゴールを右下決め打ち
    goal_row = maxrange - 1
    goal_col = maxrange - 1
    return abs((goal_row - row) + (goal_col - col))

def open_cell(row, col,):
    target = xw.Range((row, col))
    # 指定されたセルが finish 場合はゴール処理へ進ませる
    if target.color == color_finish:
        return False
    # 指定されたセルが none 以外の場合は open できないので処理を抜ける
    if target.color != color_road:
        return True
    # 確定コスト(隣接するセルの確定コストの最小値+1)を文字サイズとしてセット
    # value が存在するセル = フォントサイズを変更しているセルという前提
    tmp = []
    if xw.Range((row - 1, col)).value != None:
        tmp.append(xw.Range((row - 1, col)).api.Font.Size)
    if xw.Range((row + 1, col)).value != None:
        tmp.append(xw.Range((row + 1, col)).api.Font.Size)
    if xw.Range((row, col - 1)).value != None:
        tmp.append(xw.Range((row, col - 1)).api.Font.Size)
    if xw.Range((row, col + 1)).value != None:
        tmp.append(xw.Range((row, col + 1)).api.Font.Size)
    target.api.Font.Size = min(tmp) + 1
    # 予想コストを算出
    ret = heuristic_cost(row, col)
    # 合計コストを値としてセット
    target.options(convert=None).value = min(tmp) + 1 + ret
    # ステータスopenを背景色としてセット
    target.color = color_open
    opens.append(target)

    return True

def maizing():
    # とりあえず左上をスタート、右下をゴールとする
    #open_cell(2, 2)
    xw.Range((2,2)).color = color_start
    xw.Range((2,2)).api.Font.Size = default_fontsize + 1
    xw.Range((2,2)).options(convert=None).value = default_fontsize + 1 + heuristic_cost(2, 2)
    opens.append(xw.Range((2, 2)))

    xw.Range((maxrange - 1, maxrange - 1)).color = color_finish

    # ループ
    for i in range(2000):
        # opens = get_opens()
        # 新しい管理セルを判定する
        # 合計コストが一番小さいセルを抽出
        manager = []
        manager.append(opens[0])
        for n in opens:
            if manager[0].value > n.value:
                manager.clear
                manager.append(n)
            elif manager[0].value == n.value:
                manager.append(n)
        # 抽出したセルが複数の場合は、確定コストが最小を選択
        # <とりあえずランダムにする>
        #if len(manager) > 1:
        manager = random.choice(manager)
        # 管理セルの周辺を open する
        ret = open_cell(manager.row - 1, manager.column)
        if not ret:
            break
        ret = open_cell(manager.row + 1, manager.column)
        if not ret:
            break
        ret = open_cell(manager.row, manager.column - 1)
        if not ret:
            break
        ret = open_cell(manager.row, manager.column + 1)
        if not ret:
            break

        # 管理セルをcloseする
        manager.color = color_close
        opens.remove(manager)
    # 経路取得
    for i in range(2000):
        # 現在位置がスタート地点の場合、経路取得完了
        if manager.row == 2 and manager.column == 2:
            break
        # 現在位置の色を route 色にする
        manager.color = color_route
        # 周辺のセルをリスト化
        arounder = [xw.Range((manager.row - 1, manager.column)), xw.Range((manager.row + 1, manager.column)), xw.Range((manager.row, manager.column - 1)), xw.Range((manager.row, manager.column + 1))]
        # close セル以外は除去
        tmp = []
        for i in arounder:
            if i.color == color_close:
                tmp.append(i)
        arounder = tmp
        # 合計コストが最小のセル以上は除去
        tmp = 2000
        for i in arounder:
            if tmp > i.value:
                tmp = i.value
        for i in arounder:
            if i.value != tmp:
                arounder.remove(i)
        
        if len(arounder) == 1:
            manager = arounder[0]
        else:
            # 合計コストが同じセルが複数存在する場合は、確定コストが最小のセルを選択
            tmp = 300
            for i in arounder:
                if tmp > i.api.Font.Size:
                    tmp = i.api.Font.Size
            for i in arounder:
                if i.api.Font.Size != tmp:
                    arounder.remove(i)
            if len(arounder) == 1:
                manager = arounder[0]
            else:
                # 確定コストが同じパターンは想定外・・・
                manager = arounder[0]
    # スタート地点は close されてしまうので、最後に start 色に戻す
    xw.Range((2,2)).color = color_start


def main():
    if xw.Range((1,1)).value == finish_strings:
        # 迷路を解く
        maizing()
    else:
        make_maize()
