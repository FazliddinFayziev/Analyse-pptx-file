from charts import charts, legends
from pptx.util import Inches
from pptx.dml.color import RGBColor
from functions import calculating_all_positions, chart_size, hex_to_rgb,  get_layout_dimensions

class ChartMaker:
    def make_charts(self, slide, all_charts):
        W, H = get_layout_dimensions(slide)

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
            chart_data = chart_type[1]
            
            if "years" in chart_info:
                if len(chart_info["years"]) > 1:
                    chart_data.categories = [category for category in chart_info["years"]]
                else:
                    chart_data.categories = [category for category in chart_info["productName"]]
            else:
                chart_data.categories = [category for category in chart_info["productName"]]

            for product, series_data in zip(product_names, chart_info["data"]):
                chart_data.add_series(product, series_data)

            chart = slide.shapes.add_chart(
                chart_type[0], Inches(left), Inches(top), Inches(w), Inches(h), chart_data
            ).chart

            # COLORS
            if "colors" in chart_info:
                colors = [hex_to_rgb(color) for color in chart_info["colors"]]
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
            if "show" in chart_info and hasattr(chart.plots[0], 'has_data_labels'):
                chart.plots[0].has_data_labels = chart_info["show"]
            elif hasattr(chart.plots[0], 'has_data_labels'):
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
                legend = legends[chart_info["legend"][0]]
                chart.has_legend = True                               	
                chart.legend.position = legend
                chart.legend.include_in_layout = False
            else:
                chart.has_legend = False