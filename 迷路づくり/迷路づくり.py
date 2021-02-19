import xlwings as xw
import random

# パラメータ
# 最大範囲 (奇数のみ)
maxrange = 21
# 壁の色(白)
color_wall = (60, 180, 0)
# 道の色(濃い灰色)
color_road = (200, 200, 255)
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


def main():
    # エリアを初期化
    xw.Range((1, 1), (maxrange, maxrange)).value = ""
    xw.Range((1, 1), (maxrange, maxrange)).color = color_wall
    xw.Range((1, 1), (maxrange, maxrange)).row_height = row_height
    xw.Range((1, 1), (maxrange, maxrange)).column_width = col_width

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
    
    xw.Range((1, 1)).value = "Finish"
    xw.Range((1, 1)).select()
