import collections.abc
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from pptx.shapes.graphfrm import GraphicFrame

# create presentation
prs = Presentation()

# define layouts

slide_layout = prs.slide_layouts[6]
for i in range(10):
    if i == 0:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
    else:
        slide = prs.slides.add_slide(slide_layout)

# slide 0
title = prs.slides[0].shapes.title
subtitle = prs.slides[0].placeholders[1]
title.text = "Coding with Fazliddin"
subtitle.text = "9 January Fazliddin Fayziev"

# number of comments charts
categories = ["Product A", "Product B", "Product C", "Product D", "Product E", "Product F"]
data = [12, 24, 15, 6, 15, 8]
chart_data = CategoryChartData()
chart_data.categories = categories
chart_data.add_series('T-shirt sell from 2000 to 2005', (data))


all_data = [
    {
        "chartType": "line",
        "productName": ["Product A", "Product B", "Product C"],
        "data": [12, 24, 15],
        "chartTitle": 'T-shirt sell from 2000 to 2005'
    },
    {
        "chartType": "bar",
        "productName": ["Product A", "Product B", "Product C", "Product D", "Product E", "Product F"],
        "data": [12, 24, 15, 6, 15, 8],
        "chartTitle": 'T-shirt sell from 2000 to 2005'
    },
    {
        "chartType": "pie",
        "productName": ["Product A", "Product B", "Product C"],
        "data": [12, 24, 15],
        "chartTitle": 'T-shirt sell from 2000 to 2005'
    },
    {
        "chartType": "doughnut_exploded",
        "productName": ["Product A", "Product B", "Product C", "Product D", "Product E", "Product F"],
        "data": [12, 24, 15, 6, 15, 8],
        "chartTitle": 'T-shirt sell from 2000 to 2005'
    },
]

a_left = Inches(0.5)
b_top = Inches(0.16)
c_left = Inches(5.5)
d_top = Inches(3.80)
w_small, h_small = Inches(4), Inches(3.5)

def create_four_charts(show_values, all_charts):
    charts = {
        "pie": XL_CHART_TYPE.PIE,
        "line": XL_CHART_TYPE.LINE, 
        "doughnut": XL_CHART_TYPE.DOUGHNUT, 
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "pie_exploded": XL_CHART_TYPE.PIE_EXPLODED,
        "doughnut_exploded": XL_CHART_TYPE.DOUGHNUT_EXPLODED,
    }

    for i in range(1, 2):
        current_slide = prs.slides[i]
    
        for j, chart_info in enumerate(all_charts):
            chart_type = charts[chart_info["chartType"]]
            
            left = a_left if j % 2 == 0 else c_left
            top = b_top if j < 2 else d_top
            
            product_names = chart_info["productName"]
            chart_data = CategoryChartData()
            chart_data.categories = product_names
            chart_data.add_series(chart_info["chartTitle"], chart_info["data"])
            
            current_slide.shapes.add_chart(
                chart_type, left, top, w_small, h_small, chart_data
            ).chart.plots[0].has_data_labels = show_values

create_four_charts(True, all_data)




# IF ONLY ONE
x4, y4, cx4, cy4 = Inches(.5), Inches(.5), Inches(9), Inches(6.5)

prs.slides[2].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT, x4, y4, cx4, cy4, chart_data
)


# IF TWO
e_middle = Inches(2)

prs.slides[3].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT_EXPLODED, a_left, e_middle, w_small, h_small, chart_data
)
prs.slides[3].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT_EXPLODED, c_left, e_middle, w_small, h_small, chart_data
)


# IF TWO BUT DIFFERENT (left top, right bottom)
prs.slides[4].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT, a_left, b_top, w_small, h_small, chart_data
)
prs.slides[4].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT_EXPLODED, c_left, d_top, w_small, h_small, chart_data
)


# IF TWO BUT DIFFERENT (right top, left bottom)

prs.slides[5].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT, c_left, b_top, w_small, h_small, chart_data
)
prs.slides[5].shapes.add_chart(
    XL_CHART_TYPE.DOUGHNUT_EXPLODED, a_left, d_top, w_small, h_small, chart_data
)


# # TWO TOP

# xb, yb, cxb, cyb = Inches(0.5), Inches(0.16), Inches(4), Inches(3.5)
# xc, yc, cxc, cyc = Inches(5.5), Inches(0.16), Inches(4), Inches(3.5)

# slide6.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT, xb, yb, cxb, cyb, chart_data
# )
# slide6.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT_EXPLODED, xc, yc, cxc, cyc, chart_data
# )


# # TWO BOTTOM

# xd, yd, cxd, cyd = Inches(0.5), Inches(3.80), Inches(4), Inches(3.5)
# xe, ye, cxe, cye = Inches(5.5), Inches(3.80), Inches(4), Inches(3.5)

# slide7.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT, xd, yd, cxd, cyd, chart_data
# )
# slide7.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT_EXPLODED, xe, ye, cxe, cye, chart_data
# )


# # TWO LEFT

# xf, yf, cxf, cyf = Inches(0.5), Inches(0.16), Inches(4), Inches(3.5)
# xf0, yf0, cxf0, cyf0 = Inches(0.5), Inches(3.80), Inches(4), Inches(3.5)

# slide8.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT, xf, yf, cxf, cyf, chart_data
# )
# slide8.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT_EXPLODED, xf0, yf0, cxf0, cyf0, chart_data
# )


# # TWO RIGHT

# xf1, yf1, cxf1, cyf1 = Inches(5.5), Inches(0.16), Inches(4), Inches(3.5)
# xf2, yf2, cxf2, cyf2 = Inches(5.5), Inches(3.80), Inches(4), Inches(3.5)

# slide9.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT, xf1, yf1, cxf1, cyf1, chart_data
# )
# slide9.shapes.add_chart(
#     XL_CHART_TYPE.DOUGHNUT_EXPLODED, xf2, yf2, cxf2, cyf2, chart_data
# )

# end slide
# slide = slide2.shapes.title
# subtitle = slide2.placeholders[1]
# slide.text = "End of Presentation"
# subtitle.text = "Thank you"

prs.save('done.pptx')
print('It is done')
