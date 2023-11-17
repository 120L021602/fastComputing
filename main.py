import pandas as pd

# 导入excel文件
df = pd.read_excel('Bayesian Nash Equilibrum Q4.xlsx', sheet_name='Sheet1')

# print(df.head(10))

# 获取Player1,Player2,p1payoff,p2payoff的数据，用列表接收
player1 = df.iloc[:, 1]
player2 = df.iloc[:, 2]
p1Payoff = df.iloc[:, 3]
p2Payoff = df.iloc[:, 4]

# print(player1)
# print(player2)
# print(p1Payoff)
# print(p2Payoff)

patterns = df.iloc[:8, 5]
# print(patterns)

abyz = ['ax', 'ay', 'az', 'bx', 'by', 'bz']


def fill_in_column(player, standard):
    result = []
    for pattern in patterns:
        for str in abyz:
            if str == pattern[0] + standard[0]:
                op1 = player[abyz.index(str)]
        for str in abyz:
            if str == pattern[1] + standard[1]:
                op2 = player[abyz.index(str) + 6]
        for str in abyz:
            if str == pattern[2] + standard[2]:
                op3 = player[abyz.index(str) + 12]
        sum = op1 + op2 + op3
        result.append(sum)
    return result


def fill_in_form_and_export():
    # 填充player1的所有列
    xxxP1 = fill_in_column(p1Payoff, 'xxx')
    xxyP1 = fill_in_column(p1Payoff, 'xxy')
    xxzP1 = fill_in_column(p1Payoff, 'xxz')
    xyxP1 = fill_in_column(p1Payoff, 'xyx')
    xyyP1 = fill_in_column(p1Payoff, 'xyy')
    xyzP1 = fill_in_column(p1Payoff, 'xyz')
    xzxP1 = fill_in_column(p1Payoff, 'xzx')
    xzyP1 = fill_in_column(p1Payoff, 'xzy')
    xzzP1 = fill_in_column(p1Payoff, 'xzz')
    yxxP1 = fill_in_column(p1Payoff, 'yxx')
    yxyP1 = fill_in_column(p1Payoff, 'yxy')
    yxzP1 = fill_in_column(p1Payoff, 'yxz')
    yyxP1 = fill_in_column(p1Payoff, 'yyx')
    yyyP1 = fill_in_column(p1Payoff, 'yyy')
    yyzP1 = fill_in_column(p1Payoff, 'yyz')
    yzxP1 = fill_in_column(p1Payoff, 'yzx')
    yzyP1 = fill_in_column(p1Payoff, 'yzy')
    yzzP1 = fill_in_column(p1Payoff, 'yzz')
    zxxP1 = fill_in_column(p1Payoff, 'zxx')
    zxyP1 = fill_in_column(p1Payoff, 'zxy')
    zxzP1 = fill_in_column(p1Payoff, 'zxz')
    zyxP1 = fill_in_column(p1Payoff, 'zyx')
    zyyP1 = fill_in_column(p1Payoff, 'zyy')
    zyzP1 = fill_in_column(p1Payoff, 'zyz')
    zzxP1 = fill_in_column(p1Payoff, 'zzx')
    zzyP1 = fill_in_column(p1Payoff, 'zzy')
    zzzP1 = fill_in_column(p1Payoff, 'zzz')

    # 填充player2的所有列
    xxxP2 = fill_in_column(p2Payoff, 'xxx')
    xxyP2 = fill_in_column(p2Payoff, 'xxy')
    xxzP2 = fill_in_column(p2Payoff, 'xxz')
    xyxP2 = fill_in_column(p2Payoff, 'xyx')
    xyyP2 = fill_in_column(p2Payoff, 'xyy')
    xyzP2 = fill_in_column(p2Payoff, 'xyz')
    xzxP2 = fill_in_column(p2Payoff, 'xzx')
    xzyP2 = fill_in_column(p2Payoff, 'xzy')
    xzzP2 = fill_in_column(p2Payoff, 'xzz')
    yxxP2 = fill_in_column(p2Payoff, 'yxx')
    yxyP2 = fill_in_column(p2Payoff, 'yxy')
    yxzP2 = fill_in_column(p2Payoff, 'yxz')
    yyxP2 = fill_in_column(p2Payoff, 'yyx')
    yyyP2 = fill_in_column(p2Payoff, 'yyy')
    yyzP2 = fill_in_column(p2Payoff, 'yyz')
    yzxP2 = fill_in_column(p2Payoff, 'yzx')
    yzyP2 = fill_in_column(p2Payoff, 'yzy')
    yzzP2 = fill_in_column(p2Payoff, 'yzz')
    zxxP2 = fill_in_column(p2Payoff, 'zxx')
    zxyP2 = fill_in_column(p2Payoff, 'zxy')
    zxzP2 = fill_in_column(p2Payoff, 'zxz')
    zyxP2 = fill_in_column(p2Payoff, 'zyx')
    zyyP2 = fill_in_column(p2Payoff, 'zyy')
    zyzP2 = fill_in_column(p2Payoff, 'zyz')
    zzxP2 = fill_in_column(p2Payoff, 'zzx')
    zzyP2 = fill_in_column(p2Payoff, 'zzy')
    zzzP2 = fill_in_column(p2Payoff, 'zzz')

    data = [xxxP1, xxxP2, xxyP1, xxyP2, xxzP1, xxzP2, xyxP1, xyxP2, xyyP1, xyyP2, xyzP1, xyzP2,
            xzxP1, xzxP2, xzyP1, xzyP2, xzzP1, xzzP2, yxxP1, yxxP2, yxyP1, yxyP2, yxzP1, yxzP2,
            yyxP1, yyxP2, yyyP1, yyyP2, yyzP1, yyzP2, yzxP1, yzxP2, yzyP1, yzyP2, yzzP1, yzzP2,
            zxxP1, zxxP2, zxyP1, zxyP2, zxzP1, zxzP2, zyxP1, zyxP2, zyyP1, zyyP2, zyzP1, zyzP2,
            zzxP1, zzxP2, zzyP1, zzyP2, zzzP1, zzzP2]

    dfData = {'1': data[0], '2': data[1], '3': data[2], '4': data[3], '5': data[4],
              '6': data[5], '7': data[6], '8': data[7], '9': data[8], '10': data[9],
              '11': data[10], '12': data[11], '13': data[12], '14': data[13], '15': data[14],
              '16': data[15], '17': data[16], '18': data[17], '19': data[18], '20': data[19],
              '21': data[20], '22': data[21], '23': data[22], '24': data[23], '25': data[24],
              '26': data[25], '27': data[26], '28': data[27], '29': data[28], '30': data[29],
              '31': data[30], '32': data[31], '33': data[32], '34': data[33], '35': data[34],
              '36': data[35], '37': data[36], '38': data[37], '39': data[38], '40': data[39],
              '41': data[40], '42': data[41], '43': data[42], '44': data[43], '45': data[44],
              '46': data[45], '47': data[46], '48': data[47], '49': data[48], '50': data[49],
              '51': data[50], '52': data[51], '53': data[52], '54': data[53]}

    dfExport = pd.DataFrame(dfData)
    dfExport.to_excel('result.xlsx', index=False)


if __name__ == '__main__':
    fill_in_form_and_export()
