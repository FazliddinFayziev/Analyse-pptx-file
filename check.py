def adjust_font_size(json_data):
    for item in json_data:
        text_content = item["text_content"]
        width = item["width"]
        height = item["height"]
        
        # Calculate the ratio of text length to width and height
        length_to_width_ratio = len(text_content) / width
        length_to_height_ratio = len(text_content) / height
        
        # Set a threshold for the ratios (you can adjust this based on your requirements)
        threshold = 0.5
        
        # If the ratio exceeds the threshold, adjust the font size
        if length_to_width_ratio > threshold or length_to_height_ratio > threshold:
            # You can adjust the factor to control the font size reduction
            reduction_factor = 0.8
            item["font_size"] = reduction_factor
            
    return json_data

# Example usage
json_data = [
    {"text_content": "Elegant Education Pack for Students in Malaysia and United States", "width": 6577800, "height": 2571750, "font_size": None},
    {"text_content": "Here is where your presentation begins", "width": 6577800, "height": 458100, "font_size": None},
    {"text_content": "Done by Fazliddin and Mr Shoxrux. It is was great", "width": 6577800, "height": 458100, "font_size": None}
]

adjusted_data = adjust_font_size(json_data)
print(adjusted_data)

# Not working as expected
