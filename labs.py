import math
import pandas as pd


def laba1(N):
    print('Значения для первой лабораторной')

    R11, R21, R31, R41 = N, 2*N, 3*N, 4*N
    R12t = round(R11*R21/(R11 + R21), 3)
    R34t = round(R31*R41/(R31 + R41), 3)
    R1234t = round(R12t + R34t, 3)

    f1, f2 = 1, 10
    U = 5 + N
    I1 = float(input('Введите I1 (при f = 1kHZ): '))
    P1 = int(input('Введите P1 (при f = 1kHZ, округлите до целого): '))
    I2 = float(input('Введите I2 (при f = 10kHZ): '))
    P2 = int(input('Введите P2 (при f = 1kHZ, округлите до целого): '))

    phi1 = math.acos(P1/(U*I1))
    phi2 = math.acos(P2/(U*I2))
    Z1 = U/I1
    Z2 = U/I2
    R1 = Z1*math.cos(phi1)
    R2 = Z2*math.cos(phi2)
    X1 = Z1*math.sin(phi1)
    X2 = Z2*math.sin(phi2)
    L1 = X1/(2*math.pi*f1)
    C2 = 10**3/(2*math.pi*f2*X2)

    return [R11, R21, R31, R41, R12t, R34t, R1234t, '', '', ''], [phi1, phi2, Z1, Z2, R1, R2, X1, X2, L1, C2]


def laba2(N):


    L = (100 - 2.5 * N) * 10**(-3)
    C = (100 + 10 * N) * 10**(-6)

    f = [30, 40, 50, 60, 80, 100, 120]
    Xl, Xc = [], []
    for i in range(len(f)):
        Xl.append(round(2*math.pi*f[i]*L, 2))

    for i in range(len(f)):
        Xc.append(round(1/(2*math.pi*f[i]*C), 2))

    return Xl, Xc


def laba3(N):


    E = 10//(N**(1/4))
    R1 = 10//math.sqrt(N)
    R2 = R3 = 0.2
    L1 = (100//math.sqrt(N)) * 10**(-3)
    C1 = C2 = (100 + 10 * N) * 10**(-6)
    L2 = (round(25/math.sqrt(N), 1)) * 10**(-3)

    fPH = 1/(2*math.pi * math.sqrt(L1 * C1))

    fPT = (1/(2*math.pi * math.sqrt(L2 * C2))) * \
        (math.sqrt(L2/(C2 - R2**2)/(L2/(C2 - R3**2))))

    # При fPH:
    Xl1 = 2*math.pi*fPH*L1
    Xc1 = 1/(2*math.pi*fPH*C1)

    Z = R1 + Xl1 - Xc1
    I0 = E/Z

    Ur = I0*R1
    Ul = I0*Xl1
    Uc = I0*Xc1

    # При fPT
    Xl2 = 2*math.pi*fPT*L2
    Xc2 = 1/(2*math.pi*fPT*C2)

    Z1 = complex(R2, Xl2)
    Z2 = complex(R3, Xc2)

    Z1 = 4.07*math.e**(87.2)
    Z2 = 4.07*math.e**(-87.2)
    Z = 4.07*4.07/(0.4)

    I = E/Z

    phi1 = math.atan(Xl2/R2)
    phi2 = math.atan(Xc2/R3)

    QPH = Uc / E
    QPT = (1.02 * math.sin(-phi2)) / I22

    q = Uc/I0
    q1 = QPT / E

    delta_fPH = fPH/QPH
    delta_fPT = fPT/QPT

    E, R1, R2, R3, L1, C1, C2, L2, fPH, fPT,
    Xl1, Z, I0, Ur, Ul, Uc
    Xl2, Xc2, Z1, Z2, I, phi1, phi2, QPH, QPT, q, q1, delta_fPH, delta_fPT
    return


def laba4():
    pass


def main():
    writer = pd.ExcelWriter('./labs.xlsx', engine='xlsxwriter')
    N = int(input('Номер варианта(N): '))

    laba1_values = laba1(N)
    laba2_values = laba2(N)

    sheet_laba1 = pd.DataFrame({
        'R': ['R1', 'R2', 'R3', 'R4', 'R12(теор)', 'R34(теор)', 'R1234(теор)', '', '', ''],
        'Значения_1': list(laba1_values[0]),
        'Переменные': ['phi1', 'phi2', 'Z1', 'Z2', 'R1', 'R2', 'X1', 'X2', 'L1', 'C2'],
        'Значения_2': list(laba1_values[1])
    })

    sheet_laba2 = pd.DataFrame({
        'f': [30, 40, 50, 60, 80, 100, 120],
        'Xl': ['Xl1', 'Xl2', 'Xl3', 'Xl4', 'Xl5', 'Xl6', 'Xl7'],
        'Значения Xl': laba2_values[0],
        'Xc': ['Xc1', 'Xc2', 'Xc3', 'Xc4', 'Xc5', 'Xc6', 'Xc7'],
        'Значения Xc': laba2_values[1]
    })

    labs_sheets = {'laba1': sheet_laba1, 'laba2': sheet_laba2}
    for laba in labs_sheets.keys():
        labs_sheets[laba].to_excel(writer, sheet_name=laba, index=False)

    writer.save()
print('Значения сохранены в фале')

main()
