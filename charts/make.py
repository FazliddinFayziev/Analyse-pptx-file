from charts import charts
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from functions import calculating_all_positions, chart_size

class ChartMaker:
    def make_charts(self, slide, all_charts):
        W = 10
        H = 7.5

        for j, chart_info in enumerate(all_charts):

            product_names = chart_info["productName"]
            chart_type = charts[chart_info["chartType"]]

            # SIZE
            if type(chart_info["size"][0]) == str:
                [w, h] = chart_size(chart_info["size"][0])
            else:
                [w, h] = chart_info["size"]

            # POSITION
            if type(chart_info["position"][0]) == str:
                position = chart_info["position"][0]
                [left, top] = calculating_all_positions(position, W, H, w, h)
            else:
                [left, top] = chart_info["position"]

            # CHART
            chart_data = CategoryChartData()
            chart_data.categories = [category for category in chart_info["productName"]]

            for product, series_data in zip(product_names, chart_info["data"]):
                chart_data.add_series(product, series_data)

            chart = slide.shapes.add_chart(
                chart_type, Inches(left), Inches(top), Inches(w), Inches(h), chart_data
            ).chart

            # COLORS
            if "colors" in chart_info:
                colors = chart_info["colors"]
                if len(chart_info["data"]) == 1:
                    if "colors" in chart_info:
                        for i, point in enumerate(chart.series[0].points):
                            point.format.fill.solid()
                            point.format.fill.fore_color.rgb = RGBColor(*colors[i])
                else: 
                    for i, series in enumerate(chart.series):
                        for point in series.points:
                            if i < len(colors):
                                point.format.fill.solid()
                                point.format.fill.fore_color.rgb = RGBColor(*colors[i])

            

            # SHOW VALUE
            if "show" in chart_info:
                chart.plots[0].has_data_labels = chart_info["show"]
            else:
                chart.plots[0].has_data_labels = False

            # TITLE
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_info["chartTitle"]

            # FONT LOGIC
            if chart.has_title:
                title_format = chart.chart_title.text_frame.paragraphs[0]
                title_format.font.bold = False
                title_format.font.name = 'Arial'

            # LEGEND
            if "legend" in chart_info:
                chart.has_legend = chart_info["legend"]
            else:
                chart.has_legend = True