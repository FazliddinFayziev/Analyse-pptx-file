from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

def create_bar_chart(products, data, chart_title='Product Distribution', 
                     output_filename='BarChart.pptx', width=None, height=None, show_values=True, charts_per_slide=4, years=None):
    # Create PowerPoint presentation
    presentation = Presentation()

    # Calculate width and height based on the number of charts per slide
    if not width or not height:
        if len(data) == 1:
            width = 8
            height = 6
        elif len(data) == 2:
            width = 8
            height = 3
        else:
            width = 4
            height = 3

    # Calculate position and dimensions based on the number of charts per slide
    total_charts = len(data)
    slides_required = -(-total_charts // charts_per_slide)  # Ceiling division to get the number of slides required

    gap = 1  # Adjust this value to set the gap between charts and slides

    for slide_num in range(slides_required):
        # Add a slide
        slide_layout = presentation.slide_layouts[6]  # Use a blank slide layout
        slide = presentation.slides.add_slide(slide_layout)

        # Calculate position based on the number of charts per slide
        num_charts_on_slide = min(charts_per_slide, total_charts - slide_num * charts_per_slide)

        for i in range(num_charts_on_slide):
            if num_charts_on_slide == 3 or num_charts_on_slide == 4:
                col = i % 2
                row = i // 2
                x = Inches((col * (width + gap)) + 0.4)
                y = Inches((row * (height + gap)) + 0.2)
            elif num_charts_on_slide == 1:
                col = i % 2
                row = i // 2
                x = (Inches(10) - Inches(width)) / 2
                y = (Inches(7.5) - Inches(height)) / 2
            else:
                col = i
                x = Inches(col * (width + gap) + 0.4)
                y = Inches(2.2)

            cx, cy = Inches(width), Inches(height)

            # Add Bar chart to the slide
            chart_data = CategoryChartData()
            chart_data.categories = products
            chart_data.add_series(f'Series {i + 1}', data[slide_num * charts_per_slide + i])

            chart = slide.shapes.add_chart(
                chart_type=XL_CHART_TYPE.BAR_CLUSTERED, x=x, y=y, cx=cx, cy=cy, chart_data=chart_data
            ).chart

            # Set the chart title
            chart.has_title = True
            chart.chart_title.text_frame.text = f'{chart_title} - {years[slide_num * charts_per_slide + i]}' if years else chart_title

    # Save the PowerPoint presentation
    presentation.save(output_filename)

    print(f"Bar chart(s) created and saved in '{output_filename}'")

# Example usage
products = ['Product A', 'Product B', 'Product C', 'Product D', 'Product F']
data_single = [10, 20, 30, 10, 5, 26]
data_multiple = [
    [15, 25, 35, 25],
    [5, 15, 25, 15],
    [10, 20, 30, 10],
    [8, 18, 28, 8],
    [12, 22, 32, 12],
    [12, 22, 32, 12],
]
years = [2021, 2022, 2023, 2024, 2025, 2026]

# Single data series and no years
create_bar_chart(products, [data_single], chart_title='Single Chart', width=4, height=4, show_values=True, charts_per_slide=4, output_filename='SingleBarChart.pptx')

# Multiple data series and years
create_bar_chart(products, data_multiple, chart_title='Multiple Charts', width=4, height=4, show_values=True, charts_per_slide=4, years=years, output_filename='MultipleBarCharts.pptx')
