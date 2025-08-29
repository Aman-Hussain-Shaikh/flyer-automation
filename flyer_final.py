import os
import re
import time
import base64
from typing import List, Tuple
from pathlib import Path
import customtkinter as ctk  # All lowercase
from PIL import Image, ImageDraw, ImageFont, ImageTk
from tkinter import filedialog, messagebox, colorchooser
import pandas as pd
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import threading
import queue

class WhatsAppAutomation:
    """Handles WhatsApp Web automation using Selenium."""
    
    def __init__(self):
        self.driver = None
        self.wait = None
        self.is_logged_in = False
        
    def setup_driver(self):
        """Initialize Chrome WebDriver with more stable options."""
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            
            chrome_options = Options()
            # Add a user profile path to save login session
            user_data_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "user_data")
            chrome_options.add_argument(f"user-data-dir={user_data_dir}")
            
            # More stable options to avoid detection
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--allow-running-insecure-content")
            chrome_options.add_argument("--disable-features=VizDisplayCompositor")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Keep browser open
            chrome_options.add_experimental_option("detach", True)
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # Execute scripts to hide automation
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            self.wait = WebDriverWait(self.driver, 30) # Increased timeout
            return True
            
        except Exception as e:
            print(f"Error setting up Chrome driver: {e}")
            return False
            
    def login_to_whatsapp(self, callback=None):
        """Open WhatsApp Web and wait for user to scan QR code."""
        try:
            if not self.driver:
                if not self.setup_driver():
                    return False
            
            self.driver.get("https://web.whatsapp.com")
            
            if callback:
                callback("Please scan the QR code in your browser to login to WhatsApp Web...")
            
            # Wait for any of these elements to appear (indicating successful login)
            login_selectors = [
                "//div[@contenteditable='true'][@data-tab='3']",  # Legacy search box
                "//div[@role='textbox'][@title='Search or start new chat']", # New search box
                "//div[contains(@aria-label, 'Search')]", # Aria-label based search
                "//div[contains(@class, 'selectable-text')][@data-testid='chat-list-search']", # New structure
                "//span[@data-testid='menu']", # Menu button (indicates logged in)
            ]
            
            for _ in range(30): # 30 retries, 2 seconds each
                for selector in login_selectors:
                    try:
                        element = self.driver.find_element(By.XPATH, selector)
                        if element.is_displayed():
                            self.is_logged_in = True
                            if callback:
                                callback("Successfully logged in to WhatsApp Web!")
                            return True
                    except (NoSuchElementException, WebDriverException):
                        continue
                time.sleep(2)
            
            if callback:
                callback("Login timeout. Please try again.")
            return False
            
        except Exception as e:
            if callback:
                callback(f"Error during WhatsApp login: {e}")
            return False

    def open_chat_via_url(self, phone_number):
        """Open a chat using a direct URL. More reliable than search."""
        try:
            url = f"https://web.whatsapp.com/send?phone={phone_number}"
            self.driver.get(url)
            
            # Wait for the chat to load
            time.sleep(3)
            
            # Check for a more reliable indicator that the chat is open: the message input box
            message_input_selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@title='Type a message']",
                "//div[@role='textbox'][@title='Type a message']",
            ]
            
            for selector in message_input_selectors:
                try:
                    self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    print(f"Chat opened via URL for {phone_number}")
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue
            return False
        except Exception as e:
            print(f"Error opening chat via URL: {e}")
            return False

    def search_and_open_chat(self, contact_name_or_number):
        """
        Search for a contact and open the chat.
        This is a fallback if open_chat_via_url fails.
        """
        try:
            # Find the search box
            search_selectors = [
                "//div[@role='textbox'][@title='Search or start new chat']",
                "//div[@contenteditable='true'][@data-tab='3']",
                "//div[contains(@aria-label, 'Search')]",
            ]
            
            search_box = None
            for selector in search_selectors:
                try:
                    search_box = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not search_box:
                print("Could not find search box")
                return False
            
            # Clear and search
            search_box.clear()
            search_box.send_keys(contact_name_or_number)
            time.sleep(2)  # Wait for search results
            
            # Look for exact match in search results
            result_selectors = [
                f"//span[@title='{contact_name_or_number}']",
                f"//span[contains(text(), '{contact_name_or_number}')]",
                f"//span[contains(@title, '{contact_name_or_number}')]",
                "//div[@role='listitem'][1]"  # Fallback to the first result
            ]

            for selector in result_selectors:
                try:
                    chat_result = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    chat_result.click()
                    time.sleep(2)
                    if self._verify_chat_opened(contact_name_or_number):
                        print(f"Chat opened for {contact_name_or_number} via search.")
                        return True
                    else:
                        print(f"Verification failed for {contact_name_or_number}, trying next selector.")
                        continue
                except (TimeoutException, NoSuchElementException):
                    continue
            
            print(f"Could not open chat for: {contact_name_or_number}")
            return False
            
        except Exception as e:
            print(f"Error in search_and_open_chat: {e}")
            return False
    
    def _verify_chat_opened(self, expected_contact):
        """Verify the correct chat is opened by checking the header."""
        header_selectors = [
            f"//span[contains(@title, '{expected_contact}')]",
            f"//span[contains(text(), '{expected_contact}')]",
        ]
        
        for selector in header_selectors:
            try:
                header_element = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                if header_element.is_displayed():
                    header_text = header_element.get_attribute('title') or header_element.text
                    if expected_contact in header_text:
                        return True
            except (TimeoutException, NoSuchElementException):
                continue
        return False
        
    def send_message(self, message):
        """Send a text message with improved reliability."""
        try:
            # Multiple message box selectors
            message_selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@title='Type a message']",
                "//div[@role='textbox'][@title='Type a message']",
                "//div[contains(@class, 'selectable-text')][@contenteditable='true']",
                "//p[@class='selectable-text copyable-text'][@contenteditable='true']"
            ]
            
            message_box = None
            for selector in message_selectors:
                try:
                    message_box = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not message_box:
                print("Could not find message input box")
                return False
            
            # Type and send message
            message_box.send_keys(message)
            message_box.send_keys(Keys.ENTER)
            
            print(f"Message sent: {message[:50]}...")
            return True
            
        except Exception as e:
            print(f"Error sending message: {e}")
            return False

    def send_image(self, image_path, caption=""):
        """Send image using copy-paste method - more reliable than attachment button."""
        try:
            import pyperclip
            from PIL import Image
            import io
            
            if not os.path.exists(image_path):
                print(f"❌ Image file not found: {image_path}")
                return False
                
            print(f"📤 Attempting to send image: {os.path.basename(image_path)}")

            # Method 1: Copy image to clipboard and paste
            try:
                # Load and copy image to clipboard
                img = Image.open(image_path)
                
                # Convert to RGB if needed
                if img.mode in ('RGBA', 'LA'):
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                    img = background
                
                # Save to clipboard using win32clipboard (Windows) or alternative methods
                self._copy_image_to_clipboard(image_path)
                
                # Find the message input box
                message_input_selectors = [
                    "//div[@contenteditable='true'][@data-tab='10']",
                    "//div[@title='Type a message']",
                    "//div[@role='textbox'][@title='Type a message']",
                    "//div[contains(@class, 'selectable-text')][@contenteditable='true']",
                ]
                
                message_box = None
                for selector in message_input_selectors:
                    try:
                        message_box = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        break
                    except TimeoutException:
                        continue
                
                if not message_box:
                    print("Could not find message input box")
                    return False
                
                # Click the message box to focus it
                message_box.click()
                time.sleep(0.5)
                
                # Paste the image using Ctrl+V
                message_box.send_keys(Keys.CONTROL, 'v')
                time.sleep(3)  # Wait for image to be processed and preview to appear
                
                # Add caption if provided
                caption_box = None
                if caption:
                    try:
                        # The caption box appears after pasting image
                        caption_selectors = [
                            "//div[contains(@aria-placeholder, 'Add a caption')]",
                            "//div[@contenteditable='true' and @role='textbox']",
                            "//div[contains(@data-tab, 'caption')]",
                            "//div[contains(@class, 'selectable-text')][@contenteditable='true'][not(@data-tab='10')]"
                        ]
                        
                        for selector in caption_selectors:
                            try:
                                caption_box = self.wait.until(
                                    EC.element_to_be_clickable((By.XPATH, selector))
                                )
                                caption_box.clear()
                                caption_box.send_keys(caption)
                                print("Caption added.")
                                break
                            except TimeoutException:
                                continue
                    except Exception as e:
                        print(f"Could not add caption: {e}")
                
                # Wait a bit more for the interface to be ready
                time.sleep(1)
                
                # Look for the actual send button in the image preview interface
                send_button_found = False
                send_selectors = [
                    "//span[@data-testid='send']",
                    "//button[@data-testid='send']", 
                    "//div[@role='button'][@aria-label='Send']",
                    "//span[@data-icon='send']",
                    "//div[contains(@class, 'send')][@role='button']",
                    "//button[contains(@class, 'send')]",
                    "//span[contains(@class, 'send')][@role='button']"
                ]

                for selector in send_selectors:
                    try:
                        send_button = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        send_button.click()
                        print("Send button clicked")
                        send_button_found = True
                        break
                    except TimeoutException:
                        continue

                # If send button not found, use Enter key method
                if not send_button_found:
                    print("Send button not found, using Enter key method")
                    try:
                        if caption_box:
                            # Press Enter on the caption box
                            caption_box.send_keys(Keys.ENTER)
                            print("Enter pressed on caption box")
                        else:
                            # Press Enter on message box
                            message_box.send_keys(Keys.ENTER)
                            print("Enter pressed on message box")
                    except Exception as e:
                        print(f"Enter key method failed: {e}")
                        return False
                
                time.sleep(3)  # Wait for image to be sent
                print("Image sent successfully using copy-paste method")
                return True
                
            except Exception as e:
                print(f"Copy-paste method failed: {e}")
                # Fallback to original attachment button method
                return self._send_image_attachment_method(image_path, caption)
                
        except Exception as e:
            print(f"Error sending image: {e}")
            return False
    
    def _copy_image_to_clipboard(self, image_path):
        """Copy image to system clipboard."""
        try:
            # Windows clipboard method
            if os.name == 'nt':
                import win32clipboard
                from PIL import Image
                import io
                
                # Load image
                image = Image.open(image_path)
                
                # Convert to BMP format for clipboard
                output = io.BytesIO()
                image.convert('RGB').save(output, 'BMP')
                data = output.getvalue()[14:]  # Remove BMP header
                output.close()
                
                # Copy to clipboard
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                win32clipboard.CloseClipboard()
                
                print("Image copied to clipboard (Windows)")
                return True
                
            else:
                # For Linux/Mac - use xclip or alternative methods
                import subprocess
                
                # Try using xclip on Linux
                try:
                    subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-i', image_path], 
                                 check=True, capture_output=True)
                    print("Image copied to clipboard (Linux)")
                    return True
                except (subprocess.CalledProcessError, FileNotFoundError):
                    pass
                
                # Try using pbcopy on Mac
                try:
                    with open(image_path, 'rb') as f:
                        subprocess.run(['pbcopy'], input=f.read(), check=True)
                    print("Image copied to clipboard (Mac)")
                    return True
                except (subprocess.CalledProcessError, FileNotFoundError):
                    pass
                
                print("Could not copy image to clipboard on this system")
                return False
                
        except Exception as e:
            print(f"Error copying image to clipboard: {e}")
            return False
        
    def _send_image_attachment_method(self, image_path, caption=""):
        """Original attachment button method as fallback."""
        try:
            print("Trying original attachment button method...")
            
            # Find and click the attachment button (paperclip icon)
            attachment_button = None
            attach_selectors = [
                "//span[@data-testid='clip']",
                "//div[@title='Attach']",
                "//button[contains(@aria-label, 'Attach')]",
                "//div[@role='button'][@title='Attach']",
                "//span[@data-icon='clip']"
            ]
            
            for selector in attach_selectors:
                try:
                    attachment_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    attachment_button.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not attachment_button:
                print("Could not find attachment button")
                return False
                
            time.sleep(1)
            
            # Look for "Photos & Videos" option after clicking attach
            photo_video_selectors = [
                "//span[contains(text(), 'Photos & Videos')]",
                "//div[@title='Photos & Videos']",
                "//li[contains(., 'Photos')]"
            ]
            
            for selector in photo_video_selectors:
                try:
                    photo_option = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    photo_option.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            time.sleep(1)
            
            # Find and send keys to the hidden file input element
            file_input = None
            file_input_selectors = [
                "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']",
                "//input[@type='file'][contains(@accept, 'image')]",
                "//input[@type='file']"
            ]

            for selector in file_input_selectors:
                try:
                    file_input = self.wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                    break
                except TimeoutException:
                    continue

            if not file_input:
                print("Could not find file input for image upload.")
                return False
                
            file_input.send_keys(os.path.abspath(image_path))
            print(f"File path sent to input: {os.path.basename(image_path)}")
            
            time.sleep(3) # Wait for the image to load in the preview

            # Add caption if provided
            if caption:
                try:
                    caption_box = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@aria-placeholder, 'Add a caption')]"))
                    )
                    caption_box.send_keys(caption)
                    print("Caption added.")
                except (TimeoutException, NoSuchElementException):
                    print("Could not find caption box, sending without caption.")

            # Find and click the final send button
            send_button = None
            send_selectors = [
                "//span[@data-testid='send']",
                "//button[@data-testid='send']",
                "//div[@role='button'][@aria-label='Send']",
                "//span[@data-icon='send']"
            ]

            for selector in send_selectors:
                try:
                    send_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    send_button.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue

            if not send_button:
                print("Could not find final send button")
                return False
            
            time.sleep(2)
            print("Image sent successfully using attachment method")
            return True
            
        except Exception as e:
            print(f"Attachment method failed: {e}")
            return False
            
    def close(self):
        """Close the browser and clean up."""
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.is_logged_in = False

class ModernFlyerGeneratorApp:
    """A modern and robust flyer generation application with WhatsApp automation."""
    
    def __init__(self):
        # Configure the modern CTk appearance
        try:
            ctk.set_appearance_mode("system")
            ctk.set_default_color_theme("blue")
        except AttributeError:
            pass

        # Create main window
        self.root = ctk.CTk()
        self.root.title("🚀 Modern Flyer Generator with WhatsApp")
        self.root.geometry("1400x900")

        # Configure grid layout for the main window
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # State variables
        self.bg_image_path = ctk.StringVar()
        self.data_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.font_size = ctk.StringVar(value="36")
        self.text_color = ctk.StringVar(value="#000000")
        self.selected_font = ctk.StringVar()
        self.selected_graphic = ctk.StringVar()
        self.whatsapp_message = ctk.StringVar(value="Hi! Here's your personalized flyer.")
        
        # WhatsApp automation
        self.whatsapp_automation = WhatsAppAutomation()
        self.whatsapp_queue = queue.Queue()
        
        # Coordinate variables for text placement
        self.name_coords = {"x": 164, "y": 1437}
        self.phone_coords = {"x": 161, "y": 1509}
        
        # Variables to track dragging state
        self.dragging = False
        self.current_drag_item = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        
        # Store the original image size for coordinate calculation
        self.original_image_size = (0, 0)
        self.scale_factor = 1.0

        # Define application folders
        self.BASE_DIR = Path(__file__).parent
        self.FONT_FOLDER = self.BASE_DIR / "fonts"
        self.GRAPHICS_FOLDER = self.BASE_DIR / "graphics"
        
        # Create folders if they don't exist
        self.FONT_FOLDER.mkdir(exist_ok=True)
        self.GRAPHICS_FOLDER.mkdir(exist_ok=True)
        
        self.font_options = [f.name for f in self.FONT_FOLDER.glob("*.ttf")]

        # Set default font
        if self.font_options:
            self.selected_font.set(str(self.FONT_FOLDER / self.font_options[0]))
        else:
            self.selected_font.set("arial.ttf")

        self._setup_ui()
        self.root.bind("<Configure>", self._on_resize)
        
        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        """Handle application closing."""
        self.whatsapp_automation.close()
        self.root.destroy()

    def _load_background_image(self):
        """Opens a file dialog for the user to select a background image."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif")]
        )
        if file_path:
            self.bg_image_path.set(file_path)
            self._set_smart_default_positions()
    
    def _load_data_file(self):
        """Opens a file dialog to select a data file (CSV or XLSX)."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Data files", "*.csv *.xlsx")]
        )
        if file_path:
            self.data_path.set(file_path)
    
    def _select_output_dir(self):
        """Opens a directory dialog for the user to select an output folder."""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir.set(dir_path)
    
    def _choose_color(self):
        """Opens a color chooser dialog for selecting the text color."""
        color = colorchooser.askcolor(title="Choose text color")[1]
        if color:
            self.text_color.set(color)
            self._update_preview()
    
    def _select_graphic(self, graphic_path):
        """Sets the path for the selected graphic or icon."""
        self.selected_graphic.set(str(graphic_path))
        self._update_preview()
    
    def _on_resize(self, event):
        """Handles window resizing by updating the flyer preview."""
        if self.bg_image_path.get():
            self._update_preview()
    
    def _start_drag(self, event):
        """Starts dragging a text element."""
        self.dragging = True
        self.current_drag_item = self.preview_canvas.find_closest(event.x, event.y)[0]
        self.drag_start_x = event.x
        self.drag_start_y = event.y
    
    def _drag(self, event):
        """Handles dragging of text elements with proper bounds checking."""
        if self.dragging and self.current_drag_item:
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            
            current_x = self.preview_canvas.canvasx(event.x)
            current_y = self.preview_canvas.canvasy(event.y)
            
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if hasattr(self, 'scale_factor') and self.scale_factor > 0:
                img_width_on_canvas = int(self.original_image_size[0] * self.scale_factor)
                img_height_on_canvas = int(self.original_image_size[1] * self.scale_factor)
                
                img_left = (canvas_width - img_width_on_canvas) // 2
                img_top = (canvas_height - img_height_on_canvas) // 2
                img_right = img_left + img_width_on_canvas
                img_bottom = img_top + img_height_on_canvas
                
                constrained_x = max(img_left + 10, min(current_x, img_right - 50))
                constrained_y = max(img_top + 20, min(current_y, img_bottom - 20))
                
                bbox = self.preview_canvas.bbox(self.current_drag_item)
                if bbox:
                    current_item_x = (bbox[0] + bbox[2]) / 2
                    current_item_y = (bbox[1] + bbox[3]) / 2
                    
                    move_x = constrained_x - current_item_x
                    move_y = constrained_y - current_item_y
                    
                    self.preview_canvas.move(self.current_drag_item, move_x, move_y)
                    
                    orig_x = int((constrained_x - img_left) / self.scale_factor)
                    orig_y = int((constrained_y - img_top) / self.scale_factor)
                    
                    orig_x = max(10, min(orig_x, self.original_image_size[0] - 50))
                    orig_y = max(20, min(orig_y, self.original_image_size[1] - 20))
                    
                    if "name" in self.preview_canvas.gettags(self.current_drag_item):
                        self.name_coords = {"x": orig_x, "y": orig_y}
                    elif "phone" in self.preview_canvas.gettags(self.current_drag_item):
                        self.phone_coords = {"x": orig_x, "y": orig_y}
            
            self.drag_start_x = event.x
            self.drag_start_y = event.y
    
    def _stop_drag(self, event):
        """Stops dragging a text element."""
        self.dragging = False
        self.current_drag_item = None
    
    def _update_preview(self):
        """Updates the preview image on the canvas."""
        if not self.bg_image_path.get() or not os.path.exists(self.bg_image_path.get()):
            return
        
        if self.preview_canvas.winfo_width() <= 1:
            self.root.after(50, self._update_preview)
            return
        
        self._draw_flyer(self.preview_canvas, "Aditya Kumar", "+91 98765 43210", preview_mode=True)

    def _get_text_bounds(self, text, font):
        """Calculate the actual bounds of the text for proper positioning."""
        try:
            temp_img = Image.new('RGB', (1, 1))
            temp_draw = ImageDraw.Draw(temp_img)
            bbox = temp_draw.textbbox((0, 0), text, font=font)
            width = bbox[2] - bbox[0]
            height = bbox[3] - bbox[1]
            return width, height
        except:
            return len(text) * int(self.font_size.get()) * 0.6, int(self.font_size.get())
    
    def _constrain_coordinates(self, x, y, text_width, text_height, image_width, image_height):
        """Constrain coordinates to keep text within image bounds."""
        if x + text_width > image_width:
            x = image_width - text_width - 10
        if y + text_height > image_height:
            y = image_height - text_height - 10
        if x < 10:
            x = 10
        if y < text_height + 10:
            y = text_height + 10
        
        return int(x), int(y)

    def _set_smart_default_positions(self):
        """Sets intelligent default text positions."""
        if not self.bg_image_path.get() or not os.path.exists(self.bg_image_path.get()):
            return
        
        try:
            bg_image = Image.open(self.bg_image_path.get())
            img_width, img_height = bg_image.size
            
            name_x = max(50, int(img_width * 0.08))
            name_y = max(int(img_height * 0.75), img_height - 120)
            
            phone_x = name_x
            phone_y = min(name_y + 60, img_height - 50)
            
            font_size = int(self.font_size.get())
            
            self.name_coords = {
                "x": min(name_x, img_width - 200),
                "y": max(font_size + 10, min(name_y, img_height - font_size - 10))
            }
            
            self.phone_coords = {
                "x": min(phone_x, img_width - 200),
                "y": max(font_size + 10, min(phone_y, img_height - font_size - 10))
            }
            
            self.root.after(100, self._update_preview)
            
        except Exception as e:
            print(f"Error setting smart default positions: {e}")
    
    def _draw_flyer(self, canvas, name, phone, preview_mode=False):
        """Helper method to draw text and graphics onto an image."""
        try:
            bg_image = Image.open(self.bg_image_path.get()).convert("RGBA")
            draw = ImageDraw.Draw(bg_image)
            
            self.original_image_size = bg_image.size
            img_width, img_height = bg_image.size
            
            font_path = self.selected_font.get()
            try:
                font = ImageFont.truetype(font_path, int(self.font_size.get()))
            except:
                font = ImageFont.load_default()
            
            text_color = self.text_color.get()
            
            if preview_mode:
                canvas.delete("all")
                
                canvas_width = canvas.winfo_width()
                canvas_height = canvas.winfo_height()
                
                if canvas_width <= 1 or canvas_height <= 1:
                    canvas.update()
                    canvas_width = canvas.winfo_width()
                    canvas_height = canvas.winfo_height()
                    if canvas_width <= 1 or canvas_height <= 1:
                        canvas_width = 800
                        canvas_height = 600
                
                self.scale_factor = min(canvas_width / img_width, canvas_height / img_height)
                new_width = int(img_width * self.scale_factor)
                new_height = int(img_height * self.scale_factor)
                
                bg_image_resized = bg_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                bg_image_tk = ImageTk.PhotoImage(bg_image_resized)
                
                img_x = canvas_width // 2
                img_y = canvas_height // 2
                canvas.create_image(img_x, img_y, anchor="center", image=bg_image_tk)
                canvas._bg_image_tk = bg_image_tk
                
                img_left = img_x - new_width // 2
                img_top = img_y - new_height // 2
                
                name_canvas_x = img_left + int(self.name_coords["x"] * self.scale_factor)
                name_canvas_y = img_top + int(self.name_coords["y"] * self.scale_factor)
                phone_canvas_x = img_left + int(self.phone_coords["x"] * self.scale_factor)
                phone_canvas_y = img_top + int(self.phone_coords["y"] * self.scale_factor)
                
                scaled_font_size = max(8, int(int(self.font_size.get()) * self.scale_factor))
                
                name_text = canvas.create_text(
                    name_canvas_x, name_canvas_y,
                    text=name,
                    fill=text_color,
                    font=("Arial", scaled_font_size),
                    tags=("name", "draggable"),
                    anchor="nw"
                )
                
                phone_text = canvas.create_text(
                    phone_canvas_x, phone_canvas_y,
                    text=phone,
                    fill=text_color,
                    font=("Arial", scaled_font_size),
                    tags=("phone", "draggable"),
                    anchor="nw"
                )
                
                canvas.tag_bind("draggable", "<ButtonPress-1>", self._start_drag)
                canvas.tag_bind("draggable", "<B1-Motion>", self._drag)
                canvas.tag_bind("draggable", "<ButtonRelease-1>", self._stop_drag)
                
                canvas.create_rectangle(img_left, img_top, img_left + new_width, img_top + new_height, 
                                     outline="red", width=2, tags="image_bounds")
                
            else:
                name_width, name_height = self._get_text_bounds(name, font)
                phone_width, phone_height = self._get_text_bounds(phone, font)
                
                name_x, name_y = self._constrain_coordinates(
                    self.name_coords["x"], self.name_coords["y"],
                    name_width, name_height, img_width, img_height
                )
                
                phone_x, phone_y = self._constrain_coordinates(
                    self.phone_coords["x"], self.phone_coords["y"],
                    phone_width, phone_height, img_width, img_height
                )
                
                draw.text((name_x, name_y), name, fill=text_color, font=font)
                draw.text((phone_x, phone_y), phone, fill=text_color, font=font)

                if self.selected_graphic.get() and os.path.exists(self.selected_graphic.get()):
                    try:
                        graphic_image = Image.open(self.selected_graphic.get()).convert("RGBA")
                        graphic_size = int(self.font_size.get())
                        graphic_resized = graphic_image.resize((graphic_size, graphic_size))
                        
                        graphic_x = max(10, phone_x - graphic_size - 5)
                        graphic_y = max(10, phone_y - graphic_size // 4)
                        
                        if graphic_x + graphic_size > img_width:
                            graphic_x = img_width - graphic_size - 10
                        if graphic_y + graphic_size > img_height:
                            graphic_y = img_height - graphic_size - 10
                        
                        bg_image.paste(graphic_resized, (graphic_x, graphic_y), graphic_resized)
                    except Exception as e:
                        print(f"Error adding graphic: {e}")
                
                return bg_image

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during image processing: {e}")
            return None

    def _preview_flyer(self):
        """Trigger the preview update on button click."""
        self._update_preview()

    def _generate_flyers(self):
        """Generate flyers and save them to the output directory."""
        if not all([self.bg_image_path.get(), self.data_path.get(), self.output_dir.get()]):
            messagebox.showerror("Error", "Please select a background image, a data file, and an output directory.")
            return

        try:
            if self.data_path.get().endswith('.csv'):
                df = pd.read_csv(self.data_path.get())
            elif self.data_path.get().endswith('.xlsx'):
                df = pd.read_excel(self.data_path.get())
            else:
                messagebox.showerror("Error", "Unsupported file type. Please select a .csv or .xlsx file.")
                return

            df.columns = df.columns.str.lower().str.strip()
            
            if 'name' not in df.columns or 'number' not in df.columns:
                messagebox.showerror("Error", 
                    f"The data file must contain 'name' and 'number' columns.\n"
                    f"Found columns: {', '.join(df.columns.tolist())}")
                return

            if df.empty:
                messagebox.showerror("Error", "The data file appears to be empty.")
                return

            # Generate flyers
            for index, row in df.iterrows():
                name = str(row['name']).strip()
                phone = str(row['number']).strip()
                
                if not name or name.lower() == 'nan' or not phone or phone.lower() == 'nan':
                    continue
                
                flyer_image = self._draw_flyer(None, name, phone)
                if flyer_image:
                    sanitized_name = re.sub(r'[^a-zA-Z0-9]', '', name)
                    flyer_path = os.path.join(self.output_dir.get(), f"{sanitized_name}_flyer.png")
                    flyer_image.save(flyer_path)
            
            messagebox.showinfo("Success", "Flyers generated successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Flyer generation failed: {e}")

    def _login_whatsapp(self):
        """Login to WhatsApp Web with improved error handling."""
        def login_callback(message):
            self.root.after(0, lambda: self.status_label.configure(text=message))
        
        def login_thread():
            try:
                success = self.whatsapp_automation.login_to_whatsapp(login_callback)
                
                if success:
                    self.root.after(0, lambda: [
                        self.whatsapp_login_btn.configure(text="✅ WhatsApp Connected", state="disabled"),
                        self.whatsapp_send_btn.configure(state="normal"),
                        self.status_label.configure(text="WhatsApp connected successfully!")
                    ])
                else:
                    self.root.after(0, lambda: [
                        self.whatsapp_login_btn.configure(text="🌐 Connect WhatsApp", state="normal"),
                        self.status_label.configure(text="Failed to connect. Please try again.")
                    ])
            except Exception as e:
                self.root.after(0, lambda: [
                    self.whatsapp_login_btn.configure(text="🌐 Connect WhatsApp", state="normal"),
                    self.status_label.configure(text=f"Connection error: {str(e)}")
                ])
        
        # Disable button during login
        self.whatsapp_login_btn.configure(text="Connecting...", state="disabled")
        self.status_label.configure(text="Opening WhatsApp Web...")
        
        # Run login in separate thread to prevent UI blocking
        thread = threading.Thread(target=login_thread, daemon=True)
        thread.start()
        
    def _send_whatsapp_flyers(self):
        """Optimized WhatsApp flyer sending with faster processing."""
        if not self.whatsapp_automation.is_logged_in:
            messagebox.showerror("Error", "Please login to WhatsApp first!")
            return
        
        if not all([self.data_path.get(), self.output_dir.get()]):
            messagebox.showerror("Error", "Please generate flyers first!")
            return

        def send_messages():
            try:
                if self.data_path.get().endswith('.csv'):
                    df = pd.read_csv(self.data_path.get())
                elif self.data_path.get().endswith('.xlsx'):
                    df = pd.read_excel(self.data_path.get())
                else:
                    self.root.after(0, lambda: self.status_label.configure(text="Unsupported file type."))
                    return

                df.columns = df.columns.str.lower().str.strip()
                
                if 'name' not in df.columns or 'number' not in df.columns:
                    self.root.after(0, lambda: self.status_label.configure(text="Missing 'name' or 'number' columns."))
                    return

                total_contacts = len(df)
                sent_count = 0
                failed_contacts = []
                
                for index, row in df.iterrows():
                    name = str(row['name']).strip()
                    phone = str(row['number']).strip()
                    
                    if not name or name.lower() == 'nan' or not phone or phone.lower() == 'nan':
                        continue
                    
                    self.root.after(0, lambda n=name, sc=sent_count, tc=total_contacts: 
                        self.status_label.configure(text=f"Sending to {n} ({sc + 1}/{tc})...")
                    )
                    
                    sanitized_name = re.sub(r'[^a-zA-Z0-9_\-]', '', name)
                    flyer_path = os.path.join(self.output_dir.get(), f"{sanitized_name}_flyer.png")
                    
                    if not os.path.exists(flyer_path):
                        print(f"Flyer not found for {name}: {flyer_path}")
                        failed_contacts.append(f"{name} (flyer not found)")
                        continue
                    
                    # Try direct URL method first, then fallback to search
                    chat_opened = self.whatsapp_automation.open_chat_via_url(phone)
                    if not chat_opened:
                        print("URL method failed, falling back to search.")
                        chat_opened = self.whatsapp_automation.search_and_open_chat(name)
                    
                    if chat_opened:
                        print(f"Successfully opened chat for {name}")
                        
                        message = self.whatsapp_message.get().replace("{name}", name).replace("{phone}", phone)
                        message_sent = self.whatsapp_automation.send_message(message)
                        
                        if message_sent:
                            print(f"Message sent to {name}")
                            time.sleep(1)
                            
                            if self.whatsapp_automation.send_image(flyer_path, f"Your personalized flyer, {name}!"):
                                sent_count += 1
                                print(f"Successfully sent flyer to {name}")
                            else:
                                failed_contacts.append(f"{name} (image failed)")
                                print(f"Failed to send image to {name}")
                        else:
                            failed_contacts.append(f"{name} (message failed)")
                            print(f"Failed to send message to {name}")
                        
                        time.sleep(2)
                    else:
                        failed_contacts.append(f"{name} (contact not found)")
                        print(f"Could not find contact: {name} ({phone})")
                        time.sleep(1)
                
                success_msg = f"Completed! Sent {sent_count}/{total_contacts} flyers."
                if failed_contacts:
                    failed_msg = f"\n\nFailed contacts:\n" + "\n".join(failed_contacts[:10])
                    if len(failed_contacts) > 10:
                        failed_msg += f"\n... and {len(failed_contacts) - 10} more"
                    success_msg += failed_msg
                
                self.root.after(0, lambda: self.status_label.configure(text=success_msg))
                self.root.after(0, lambda: messagebox.showinfo("WhatsApp Sending Complete", 
                    f"Successfully sent: {sent_count}/{total_contacts} flyers\n"
                    f"Failed: {len(failed_contacts)} contacts"
                ))
                
            except Exception as e:
                error_msg = f"Error during WhatsApp sending: {e}"
                self.root.after(0, lambda: self.status_label.configure(text=error_msg))
                self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            
            finally:
                self.root.after(0, lambda: self.whatsapp_send_btn.configure(
                    text="📱 Send Flyers", state="normal"
                ))
        
        self.whatsapp_send_btn.configure(text="Sending...", state="disabled")
        thread = threading.Thread(target=send_messages, daemon=True)
        thread.start()

    def _setup_ui(self):
        """Sets up the graphical user interface elements."""
        self.sidebar_frame = ctk.CTkFrame(self.root, width=300)
        self.sidebar_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        self.title_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Flyer Generator", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=20)

        self.tab_view = ctk.CTkTabview(self.sidebar_frame)
        self.tab_view.pack(padx=10, pady=10, fill="both", expand=True)

        tabs = [
            ("Image", self._create_image_tab),
            ("Text", self._create_text_tab),
            ("Graphics", self._create_graphics_tab),
            ("WhatsApp", self._create_whatsapp_tab)
        ]

        for name, setup_func in tabs:
            tab = self.tab_view.add(name)
            setup_func(tab)

        self.action_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.action_frame.pack(pady=10, fill="x")

        self.button_row1 = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        self.button_row1.pack(fill="x", pady=5)

        ctk.CTkButton(
            self.button_row1, 
            text="🔍 Preview", 
            command=self._preview_flyer,
            fg_color="green",
            hover_color="dark green",
            width=140
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            self.button_row1, 
            text="🎨 Generate", 
            command=self._generate_flyers,
            fg_color="blue",
            hover_color="dark blue",
            width=140
        ).pack(side="right", padx=5)

        self.whatsapp_buttons_frame = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        self.whatsapp_buttons_frame.pack(fill="x", pady=5)

        self.whatsapp_login_btn = ctk.CTkButton(
            self.whatsapp_buttons_frame,
            text="🌐 Connect WhatsApp",
            command=self._login_whatsapp,
            fg_color="green",
            hover_color="dark green",
            width=140
        )
        self.whatsapp_login_btn.pack(side="left", padx=5)

        self.whatsapp_send_btn = ctk.CTkButton(
            self.whatsapp_buttons_frame,
            text="📱 Send Flyers",
            command=self._send_whatsapp_flyers,
            fg_color="orange",
            hover_color="dark orange",
            state="disabled",
            width=140
        )
        self.whatsapp_send_btn.pack(side="right", padx=5)

        self.status_label = ctk.CTkLabel(
            self.action_frame,
            text="Ready to generate flyers",
            text_color="gray"
        )
        self.status_label.pack(pady=5)

        self.preview_frame = ctk.CTkFrame(self.root)
        self.preview_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)

        self.preview_canvas = ctk.CTkCanvas(
            self.preview_frame, 
            bg="white", 
            highlightthickness=0
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        
        self.instructions_label = ctk.CTkLabel(
            self.preview_frame,
            text="Drag the text elements to position them on the flyer",
            text_color="gray"
        )
        self.instructions_label.grid(row=1, column=0, pady=5)

    def _create_image_tab(self, master):
        """Sets up the controls for image and data file selection."""
        controls = [
            ("Background Image", self.bg_image_path, self._load_background_image),
            ("Data File (.csv/.xlsx)", self.data_path, self._load_data_file),
            ("Output Directory", self.output_dir, self._select_output_dir)
        ]

        for label, var, command in controls:
            frame = ctk.CTkFrame(master, fg_color="transparent")
            frame.pack(pady=10, fill="x")
            
            ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
            
            entry_frame = ctk.CTkFrame(frame, fg_color="transparent")
            entry_frame.pack(fill="x")
            
            entry = ctk.CTkEntry(entry_frame, textvariable=var)
            entry.pack(side="left", expand=True, fill="x", padx=(0, 5))
            
            ctk.CTkButton(
                entry_frame, 
                text="Browse", 
                command=command, 
                width=80
            ).pack(side="right")

    def _create_text_tab(self, master):
        """Sets up the controls for text styling."""
        font_frame = ctk.CTkFrame(master, fg_color="transparent")
        font_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(font_frame, text="Font", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        if self.font_options:
            font_menu = ctk.CTkComboBox(
                font_frame, 
                values=self.font_options,
                variable=self.selected_font
            )
            font_menu.pack(fill="x")
            font_menu.bind("<<ComboboxSelected>>", lambda e: self._update_preview())
        else:
            ctk.CTkLabel(font_frame, text="No fonts found in 'fonts' folder", text_color="red").pack()

        size_frame = ctk.CTkFrame(master, fg_color="transparent")
        size_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(size_frame, text="Font Size", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        size_entry = ctk.CTkEntry(size_frame, textvariable=self.font_size)
        size_entry.pack(fill="x")
        size_entry.bind("<Return>", lambda e: self._update_preview())

        color_frame = ctk.CTkFrame(master, fg_color="transparent")
        color_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(color_frame, text="Text Color", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        color_entry_frame = ctk.CTkFrame(color_frame, fg_color="transparent")
        color_entry_frame.pack(fill="x")
        
        color_entry = ctk.CTkEntry(color_entry_frame, textvariable=self.text_color)
        color_entry.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        ctk.CTkButton(
            color_entry_frame, 
            text="Choose", 
            command=self._choose_color, 
            width=80
        ).pack(side="right")

    def _create_graphics_tab(self, master):
        """Sets up the controls for selecting graphics/icons."""
        ctk.CTkLabel(master, text="Graphics/Icons", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=10)
        
        graphics_files = list(self.GRAPHICS_FOLDER.glob("*"))
        
        if graphics_files:
            graphics_grid = ctk.CTkScrollableFrame(master)
            graphics_grid.pack(fill="both", expand=True, padx=5, pady=5)

            for graphic in graphics_files:
                try:
                    img = ctk.CTkImage(
                        Image.open(graphic).resize((80, 80)), 
                        size=(80, 80)
                    )
                    btn = ctk.CTkButton(
                        graphics_grid, 
                        image=img, 
                        text=graphic.stem[:10], 
                        compound="top",
                        command=lambda g=graphic: self._select_graphic(g),
                        width=100,
                        height=100
                    )
                    btn.pack(pady=5)
                except Exception as e:
                    print(f"Error loading graphic {graphic}: {e}")
        else:
            ctk.CTkLabel(
                master, 
                text="No graphics found.\nAdd images to the 'graphics' folder.",
                text_color="gray"
            ).pack(pady=20)

    def _create_whatsapp_tab(self, master):
        """Sets up WhatsApp automation controls."""
        instructions = ctk.CTkTextbox(master, height=120, wrap="word")
        instructions.pack(pady=10, fill="x")
        instructions.insert("0.0", 
            "WhatsApp Automation Instructions:\n\n"
            "1. Click 'Connect WhatsApp' to open WhatsApp Web\n"
            "2. Scan the QR code with your phone\n"
            "3. Generate flyers first\n"
            "4. Click 'Send Flyers' to automatically send to all contacts\n\n"
            "Note: Your phone must stay connected to internet during sending."
        )
        instructions.configure(state="disabled")

        message_frame = ctk.CTkFrame(master, fg_color="transparent")
        message_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(
            message_frame, 
            text="Custom Message", 
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", pady=2)
        
        message_entry = ctk.CTkTextbox(message_frame, height=60)
        message_entry.pack(fill="x")
        message_entry.insert("0.0", self.whatsapp_message.get())
        
        def update_message():
            self.whatsapp_message.set(message_entry.get("0.0", "end-1c"))
        
        message_entry.bind("<KeyRelease>", lambda e: update_message())

        tips_frame = ctk.CTkFrame(master, fg_color="transparent")
        tips_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(tips_frame, text="💡 Tips:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(
            tips_frame, 
            text="• Use {name} in message for personalization\n"
                 "• Test with a few contacts first\n"
                 "• Wait 3-5 seconds between messages\n"
                 "• Keep your phone connected",
            text_color="gray",
            justify="left"
        ).pack(anchor="w", pady=5)

        warning_frame = ctk.CTkFrame(master, fg_color="#ffebcc")
        warning_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(
            warning_frame,
            text="⚠️ Important: Use responsibly and comply with WhatsApp's terms of service.\n"
                 "Sending too many messages quickly may result in temporary restrictions.",
            text_color="#cc6600",
            wraplength=250,
            justify="left"
        ).pack(pady=10)

def check_dependencies():
    """Check if all required packages are installed."""
    required_packages = {
        'selenium': 'selenium',
        'pandas': 'pandas', 
        'openpyxl': 'openpyxl',
        'customtkinter': 'customtkinter',
        'PIL': 'Pillow',
        'webdriver_manager': 'webdriver-manager'
    }
    
    missing = []
    
    for import_name, pip_name in required_packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)
    
    if missing:
        print("Missing required packages:")
        for package in missing:
            print(f"  pip install {package}")
        return False
    
    return True

def main():
    """Initialize and run the main application."""
    if not check_dependencies():
        input("Press Enter to exit...")
        return
    
    print("Starting Modern Flyer Generator with WhatsApp Automation...")
    print("Make sure ChromeDriver is installed and in your PATH!")
    
    app = ModernFlyerGeneratorApp()
    app.root.mainloop()

if __name__ == "__main__":
    main()
