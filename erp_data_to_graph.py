import matplotlib.pyplot as plt
import matplotlib


def create_graph(title,storage_price,sales_avg):
    matplotlib.rcParams["font.family"] = "Malgun Gothic"
    matplotlib.rcParams["font.size"] = 15 # 글자 크기
    matplotlib.rcParams["axes.unicode_minus"] = False # 한글 폰트 사용 시, 마이너스 글자가 깨지는 현상 해결

    fig, axs = plt.subplots(1,2, figsize=(15,10))
    fig.suptitle(title)


    labels = ['','월 평균 매출(12개월)','원자재 재고 금액',' ']
    values = [0,sales_avg,storage_price,0]

    pcts = [0.15,0.1,0.05]

    red_card = [int(values[1]*pcts[0])] * 4
    yellow_card = [int(values[1]*pcts[1])] * 4
    green_card = [int(values[1]*pcts[2])] * 4

    colors = ["grey","blue"]

    axs1 = axs[0]
    axs1.set_title("월 평균 매출(12개월)")

    axs1.bar(labels,values, color = colors, width = 0.8)

    axs1.text(1,values[1]+700,values[1],ha="center")
    axs1.text(2,values[2]+700,values[2],ha="center")

    axs1.plot(labels, red_card, label = "Red Card", color='red', linestyle='dashed')
    axs1.plot(labels, yellow_card, label = "Yellow Card", color='orange', linestyle='dashed')
    axs1.plot(labels, green_card, label = "Green Card", color='green', linestyle='dashed')

    axs1.text(3,red_card[0]-500,red_card[0],ha="left")
    axs1.text(3,yellow_card[0]-500,yellow_card[0],ha="left")
    axs1.text(3,green_card[0]-500,green_card[0],ha="left")

    axs1.set(ylabel = 'USD') # y축 이름
    axs1.set_xticklabels(labels,rotation=25)

    axs1.legend() # 범례

    ##########################

    labels = ['','원자재 재고 금액',' ']
    values = [0,storage_price,0]

    pcts = [0.15,0.1,0.05]

    red_card = red_card[1:]
    yellow_card = yellow_card[1:]
    green_card = green_card[1:]

    colors = ["grey","blue"]

    axs2 = axs[1]
    axs2.set_title("원자재 금액")

    axs2.bar(labels,values, color = colors[0])

    axs2.text(1,values[1]+700,values[1],ha="center")

    axs2.plot(labels, red_card, label = "Red Card", color='red', linestyle='dashed')
    axs2.plot(labels, yellow_card, label = "Yellow Card", color='orange', linestyle='dashed')
    axs2.plot(labels, green_card, label = "Green Card", color='green', linestyle='dashed')

    axs2.text(2,red_card[0],red_card[0],ha="left")
    axs2.text(2,yellow_card[0],yellow_card[0],ha="left")
    axs2.text(2,green_card[0],green_card[0],ha="left")

    axs2.set(ylabel = 'USD') # y축 이름
    # axs2.legend() # 범례

    plt.savefig(f'{title}.png')

