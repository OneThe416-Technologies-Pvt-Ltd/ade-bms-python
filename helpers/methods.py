#method.py

from PIL import Image, ImageTk, ImageOps

def load_and_resize_image(image_path, size, border_width=0, border_color=None):
    try:
        # Load and convert image
        original_image = Image.open(image_path).convert("RGBA")
        if border_width > 0 and border_color:
            # Add border to image if specified
            original_image = ImageOps.expand(original_image, border=border_width, fill=border_color)
        
        # Resize the image
        resized_image = original_image.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(resized_image)
    except Exception as e:
        print(f"Error loading or resizing image: {e}")
        return None

def create_image_button(canvas, image_tk, position, click_handler):
    if image_tk is not None:
        btn_item = canvas.create_image(position, image=image_tk)
        canvas.tag_bind(btn_item, '<Button-1>', click_handler)
        canvas.tag_bind(btn_item, '<Enter>', lambda event: canvas.config(cursor="hand2"))
        canvas.tag_bind(btn_item, '<Leave>', lambda event: canvas.config(cursor=""))
        return btn_item
    else:
        print(f"Error: Image for button at position {position} is None")
        return None

