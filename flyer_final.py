import os
import re
import time
import base64
from typing import List, Tuple
from pathlib import Path
import customtkinter as ctk
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
        """Initialize Chrome WebDriver with optimized options for faster performance."""
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            
            chrome_options = Options()
            # User profile for persistent login
            user_data_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "user_data")
            chrome_options.add_argument(f"user-data-dir={user_data_dir}")
            
            # Performance optimizations
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--disable-features=VizDisplayCompositor")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-images")  # Faster loading
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Keep browser open
            chrome_options.add_experimental_option("detach", True)
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # Execute scripts to hide automation
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            self.wait = WebDriverWait(self.driver, 20)  # Wait for up to 20 seconds
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
            
            # Wait for login indicators
            login_selectors = [
                "//div[@contenteditable='true'][@data-tab='3']",
                "//div[@role='textbox'][@title='Search or start new chat']",
                "//div[contains(@aria-label, 'Search')]",
                "//div[contains(@class, 'selectable-text')][@data-testid='chat-list-search']",
                "//span[@data-testid='menu']",
            ]
            
            for _ in range(30):
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
        """Open a chat using direct URL - fastest method."""
        try:
            url = f"https://web.whatsapp.com/send?phone={phone_number}"
            self.driver.get(url)
            
            # Use WebDriverWait to confirm chat is open, which is more robust than a fixed sleep
            message_input_selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@title='Type a message']",
                "//div[@role='textbox'][@title='Type a message']",
            ]
            
            for selector in message_input_selectors:
                try:
                    self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    print(f"Chat opened for {phone_number}")
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue
            return False
        except Exception as e:
            print(f"Error opening chat via URL: {e}")
            return False

    def search_and_open_chat(self, contact_name_or_number):
        """Fallback search method with reduced timeouts."""
        try:
            search_selectors = [
                "//div[@role='textbox'][@title='Search or start new chat']",
                "//div[@contenteditable='true'][@data-tab='3']",
                "//div[contains(@aria-label, 'Search')]",
            ]
            
            search_box = None
            for selector in search_selectors:
                try:
                    search_box = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not search_box:
                return False
            
            search_box.clear()
            search_box.send_keys(contact_name_or_number)
            
            # Dynamic wait for search results
            result_selectors = [
                f"//span[@title='{contact_name_or_number}']",
                f"//span[contains(text(), '{contact_name_or_number}')]",
                "//div[@role='listitem'][1]"
            ]

            for selector in result_selectors:
                try:
                    chat_result = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    chat_result.click()
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue
            
            return False
            
        except Exception as e:
            print(f"Error in search_and_open_chat: {e}")
            return False
        
    def send_message(self, message):
        """Send a text message with improved delivery confirmation."""
        try:
            message_selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@title='Type a message']",
                "//div[@role='textbox'][@title='Type a message']",
            ]
            
            message_box = None
            for selector in message_selectors:
                try:
                    message_box = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not message_box:
                print("Could not find message input box")
                return False
            
            # Clear any existing text and type the new message
            message_box.clear()
            message_box.click()  # Ensure focus
            time.sleep(0.3)  # Small delay to ensure focus
            
            message_box.send_keys(message)
            message_box.send_keys(Keys.ENTER)
            
            # Wait for message to be sent by checking for delivery indicators
            try:
                # Wait for the message to appear in chat (usually takes 1-3 seconds)
                WebDriverWait(self.driver, 10).until(
                    lambda driver: driver.execute_script(
                        "return document.querySelector('[data-testid=\"msg-container\"]') !== null"
                    )
                )
                
                # Additional wait to ensure message is fully processed
                time.sleep(2)
                
                # Verify message input is ready for next message
                WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, message_selectors[0]))
                )
                
            except TimeoutException:
                print("Warning: Could not confirm message delivery")
                time.sleep(3)  # Fallback wait time
            
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
                        message_box = WebDriverWait(self.driver, 5).until(
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
                time.sleep(0.5) # Small sleep to ensure focus
                
                # Paste the image using Ctrl+V
                message_box.send_keys(Keys.CONTROL, 'v')
                
                # --- START OF CRITICAL IMPROVEMENT FOR SPEED ---
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

                # Wait for the send button to be clickable inside the image preview
                for selector in send_selectors:
                    try:
                        send_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        # Add caption if provided AFTER the preview appears
                        if caption:
                            try:
                                caption_box = WebDriverWait(self.driver, 2).until(
                                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@aria-placeholder, 'Add a caption')]"))
                                )
                                caption_box.clear()
                                caption_box.send_keys(caption)
                                print("Caption added.")
                            except (TimeoutException, NoSuchElementException):
                                print("Could not find caption box.")
                        
                        send_button.click()
                        print("Send button clicked")
                        send_button_found = True
                        break
                    except TimeoutException:
                        continue

                if not send_button_found:
                    print("Send button not found. Attempting to use Enter key.")
                    # Fallback to pressing Enter on the message box if a send button wasn't found
                    target_box = caption_box if caption and 'caption_box' in locals() else message_box
                    if target_box:
                        target_box.send_keys(Keys.ENTER)
                        send_button_found = True
                
                if send_button_found:
                    # Final dynamic wait to confirm message is sent
                    # The message input box should become clickable again.
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, message_input_selectors[0])))
                    print("Image sent successfully using copy-paste method")
                    return True
                else:
                    print("Failed to send image.")
                    return False
                # --- END OF CRITICAL IMPROVEMENT FOR SPEED ---
                
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
            if os.name == 'nt':
                import win32clipboard
                from PIL import Image
                import io
                
                image = Image.open(image_path)
                output = io.BytesIO()
                image.convert('RGB').save(output, 'BMP')
                data = output.getvalue()[14:]
                output.close()
                
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                win32clipboard.CloseClipboard()
                
                return True
            else:
                # Use a command-line tool for Linux/macOS
                import subprocess
                try:
                    # Try xclip for Linux
                    subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-i', image_path],  
                                   check=True, capture_output=True)
                    return True
                except (subprocess.CalledProcessError, FileNotFoundError):
                    try:
                        # Try pbcopy for macOS
                        with open(image_path, 'rb') as f:
                            subprocess.run(['pbcopy'], input=f.read(), check=True)
                        return True
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        return False
        except Exception as e:
            print(f"Error copying image to clipboard: {e}")
            return False
            
    def _send_image_attachment_method(self, image_path, caption=""):
        """Fallback attachment method with reduced timeouts."""
        try:
            attach_selectors = [
                "//span[@data-testid='clip']",
                "//div[@title='Attach']",
            ]
            
            attachment_button = None
            for selector in attach_selectors:
                try:
                    attachment_button = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    attachment_button.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not attachment_button:
                return False
                
            photo_video_selectors = [
                "//span[contains(text(), 'Photos & Videos')]",
                "//div[@title='Photos & Videos']",
            ]
            
            for selector in photo_video_selectors:
                try:
                    photo_option = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    photo_option.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            file_input_selectors = [
                "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']",
                "//input[@type='file'][contains(@accept, 'image')]",
            ]

            file_input = None
            for selector in file_input_selectors:
                try:
                    file_input = WebDriverWait(self.driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue

            if not file_input:
                return False
                
            file_input.send_keys(os.path.abspath(image_path))
            
            # Use dynamic wait instead of fixed sleep
            try:
                caption_box = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@aria-placeholder, 'Add a caption')]"))
                )
                if caption:
                    caption_box.send_keys(caption)
            except (TimeoutException, NoSuchElementException):
                print("Could not find caption box.")

            # Send button
            send_selectors = [
                "//span[@data-testid='send']",
                "//button[@data-testid='send']",
            ]

            for selector in send_selectors:
                try:
                    send_button = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    send_button.click()
                    break
                except (TimeoutException, NoSuchElementException):
                    continue

            # Final check to see if the message input box is available again
            message_input_selectors = [
                "//div[@contenteditable='true'][@data-tab='10']",
                "//div[@title='Type a message']",
                "//div[@role='textbox'][@title='Type a message']",
            ]
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, message_input_selectors[0])))

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
    """Enhanced flyer generation application with coordinate-based positioning and improved styling."""
    
    def __init__(self):
        try:
            ctk.set_appearance_mode("system")
            ctk.set_default_color_theme("blue")
        except AttributeError:
            pass

        # Create main window
        self.root = ctk.CTk()
        self.root.title("Enhanced Flyer Generator with WhatsApp")
        self.root.geometry("1500x950")

        # Configure grid layout
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
        
        # Enhanced WhatsApp variables
        self.whatsapp_message = ctk.StringVar(value="Hi {name}! Here's your personalized flyer.")
        self.use_custom_message = ctk.BooleanVar(value=False)
        self.image_caption = ctk.StringVar(value="Your personalized flyer!")
        self.use_custom_caption = ctk.BooleanVar(value=False)
        
        # Text styling variables
        self.text_bold = ctk.BooleanVar(value=False)
        self.text_italic = ctk.BooleanVar(value=False)
        self.text_underline = ctk.BooleanVar(value=False)
        self.text_shadow = ctk.BooleanVar(value=False)
        self.shadow_color = ctk.StringVar(value="#808080")
        
        self.name_x = ctk.StringVar(value="500")
        self.name_y = ctk.StringVar(value="1900")
        self.phone_x = ctk.StringVar(value="490")
        self.phone_y = ctk.StringVar(value="1970")
        
        # WhatsApp automation
        self.whatsapp_automation = WhatsAppAutomation()
        
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
        
        # Load all font files
        self.font_options = []
        for ext in ['*.ttf', '*.TTF', '*.otf', '*.OTF']:
            self.font_options.extend([f.name for f in self.FONT_FOLDER.glob(ext)])
        
        # Add system fonts fallback
        if not self.font_options:
            self.font_options = ["arial.ttf", "calibri.ttf", "times.ttf", "helvetica.ttf"]

        # Set default font
        if self.font_options:
            self.selected_font.set(str(self.FONT_FOLDER / self.font_options[0]))
        else:
            self.selected_font.set("arial.ttf")

        self._setup_ui()
        self.root.bind("<Configure>", self._on_resize)
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        """Handle application closing."""
        self.whatsapp_automation.close()
        self.root.destroy()

    def _load_background_image(self):
        """Opens a file dialog for the user to select a background image."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp")]
        )
        if file_path:
            self.bg_image_path.set(file_path)
            self._update_preview()
    
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
    
    def _choose_shadow_color(self):
        """Opens a color chooser dialog for selecting the shadow color."""
        color = colorchooser.askcolor(title="Choose shadow color")[1]
        if color:
            self.shadow_color.set(color)
            self._update_preview()
    
    def _select_graphic(self, graphic_path):
        """Sets the path for the selected graphic or icon."""
        self.selected_graphic.set(str(graphic_path))
        self._update_preview()
    
    def _on_resize(self, event):
        """Handles window resizing by updating the flyer preview."""
        if self.bg_image_path.get():
            self._update_preview()
    
    def _update_coordinates(self):
        """Update preview when coordinates change."""
        self._update_preview()
    
    def _update_preview(self):
        """
        Updates the preview image on the canvas by drawing a full preview
        image in memory using PIL and then displaying it.
        """
        if not self.bg_image_path.get() or not os.path.exists(self.bg_image_path.get()):
            self.status_label.configure(text="Please select a background image.", text_color="orange")
            return

        try:
            # Create a temporary image with the preview data, just like the final flyer
            preview_image = self._draw_flyer("Coreprix", "+91 90000 XXXXX")
            if not preview_image:
                self.status_label.configure(text="Failed to generate preview image.", text_color="red")
                return

            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()

            if canvas_width <= 1 or canvas_height <= 1:
                self.root.after(50, self._update_preview)
                return

            img_width, img_height = preview_image.size
            self.scale_factor = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * self.scale_factor)
            new_height = int(img_height * self.scale_factor)
            
            # Resize the pre-rendered Pillow image to fit the canvas
            resized_preview_image = preview_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert to a format Tkinter can display
            self.preview_image_tk = ImageTk.PhotoImage(resized_preview_image)
            
            # Clear the old canvas contents and display the new image
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(
                canvas_width // 2, 
                canvas_height // 2, 
                anchor="center", 
                image=self.preview_image_tk
            )
            
            self.status_label.configure(text="Preview updated successfully", text_color="gray")
            
        except Exception as e:
            messagebox.showerror("Preview Error", f"An error occurred while updating preview: {e}")
            self.status_label.configure(text="Error updating preview", text_color="red")
    
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
            # Fallback for font errors
            return len(text) * int(self.font_size.get()) * 0.6, int(self.font_size.get())

    def _apply_text_effects(self, draw, text, position, font, color):
        """Apply text effects like shadow, bold, underline, etc."""
        x, y = position
        
        # Apply shadow effect
        if self.text_shadow.get():
            shadow_offset = max(2, int(self.font_size.get()) // 15)
            draw.text((x + shadow_offset, y + shadow_offset), text, 
                      fill=self.shadow_color.get(), font=font)
        
        # Simulate bold by drawing text multiple times with slight offsets
        if self.text_bold.get():
            for dx in range(1, 3):
                for dy in range(1, 3):
                    draw.text((x + dx, y + dy), text, fill=color, font=font)
        
        # Draw main text
        draw.text((x, y), text, fill=color, font=font)
        
        # Apply underline effect
        if self.text_underline.get():
            try:
                bbox = draw.textbbox((x, y), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                underline_y = y + text_height + 2
                underline_thickness = max(1, int(self.font_size.get()) // 20)
                
                for i in range(underline_thickness):
                    draw.line([(x, underline_y + i), (x + text_width, underline_y + i)], 
                              fill=color, width=1)
            except:
                pass
    
    def _create_italic_font_image(self, text, font, color):
        """Create an italicized version of text by skewing the image."""
        try:
            # Create temporary image for text
            temp_img = Image.new('RGBA', (1000, 200), (0, 0, 0, 0))
            temp_draw = ImageDraw.Draw(temp_img)
            temp_draw.text((50, 50), text, fill=color, font=font)
            
            # Apply italic transformation (skew)
            if self.text_italic.get():
                from PIL import ImageTransform
                skew_factor = 0.2
                coeffs = (1, skew_factor, 0, 0, 1, 0)
                temp_img = temp_img.transform(temp_img.size, ImageTransform.AFFINE, coeffs)
            
            return temp_img
        except:
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

            total_count = 0
            for index, row in df.iterrows():
                name = str(row['name']).strip()
                phone = str(row['number']).strip()
                
                if not name or name.lower() == 'nan' or not phone or phone.lower() == 'nan':
                    continue
                
                flyer_image = self._draw_flyer(name, phone)
                if flyer_image:
                    sanitized_name = re.sub(r'[^a-zA-Z0-9]', '', name)
                    flyer_path = os.path.join(self.output_dir.get(), f"{sanitized_name}_flyer.png")
                    flyer_image.save(flyer_path)
                    total_count += 1
            
            messagebox.showinfo("Success", f"Generated {total_count} flyers successfully!")

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
        
        self.whatsapp_login_btn.configure(text="Connecting...", state="disabled")
        self.status_label.configure(text="Opening WhatsApp Web...")
        
        thread = threading.Thread(target=login_thread, daemon=True)
        thread.start()
        
    def _send_whatsapp_flyers(self):
        """Optimized WhatsApp flyer sending with proper message delivery timing."""
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
                        failed_contacts.append(f"{name} (flyer not found)")
                        time.sleep(1)
                        continue
                    
                    chat_opened = self.whatsapp_automation.open_chat_via_url(phone)
                    if not chat_opened:
                        chat_opened = self.whatsapp_automation.search_and_open_chat(name)
                    
                    if chat_opened:
                        if self.use_custom_message.get():
                            message = self.whatsapp_message.get().replace("{name}", name).replace("{phone}", phone)
                            message_sent = self.whatsapp_automation.send_message(message)
                            if not message_sent:
                                failed_contacts.append(f"{name} (message failed)")
                                time.sleep(1)
                                continue
                            # No additional sleep needed here since send_message() already waits 2 seconds
                        
                        caption = ""
                        if self.use_custom_caption.get():
                            caption = self.image_caption.get().replace("{name}", name).replace("{phone}", phone)
                        
                        if self.whatsapp_automation.send_image(flyer_path, caption):
                            sent_count += 1
                            print(f"Successfully sent flyer to {name}")
                            # Wait 2 seconds after image is sent to ensure proper delivery
                            time.sleep(2)
                        else:
                            failed_contacts.append(f"{name} (image failed)")
                            time.sleep(1)
                            
                    else:
                        failed_contacts.append(f"{name} (contact not found)")
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

    def _draw_flyer(self, name, phone):
        """
        Enhanced flyer drawing with coordinate-based positioning and styling.
        Draws directly to a Pillow Image object.
        """
        try:
            bg_image = Image.open(self.bg_image_path.get()).convert("RGBA")
            draw = ImageDraw.Draw(bg_image)
            
            self.original_image_size = bg_image.size
            
            try:
                font_size = int(self.font_size.get() or 36)
                name_x = int(float(self.name_x.get() or 0))
                name_y = int(float(self.name_y.get() or 0))
                phone_x = int(float(self.phone_x.get() or 0))
                phone_y = int(float(self.phone_y.get() or 0))
            except ValueError:
                messagebox.showerror("Input Error", "Please enter valid numeric values for positions and font size.")
                return None

            font_path = self.selected_font.get()
            try:
                font_file = Path(font_path)
                if not font_file.exists():
                    font_file = self.FONT_FOLDER / font_file.name
                
                font = ImageFont.truetype(str(font_file), font_size)
            except Exception as e:
                print(f"Font loading error: {e}. Using default font.")
                font = ImageFont.load_default()
            
            text_color = self.text_color.get()
            
            # Apply styling to the text and draw it on the image
            name_pos = (name_x, name_y)
            phone_pos = (phone_x, phone_y)
            
            self._apply_text_effects(draw, name, name_pos, font, text_color)
            self._apply_text_effects(draw, phone, phone_pos, font, text_color)

            # Add graphic if selected
            if self.selected_graphic.get() and os.path.exists(self.selected_graphic.get()):
                try:
                    graphic_image = Image.open(self.selected_graphic.get()).convert("RGBA")
                    graphic_size = font_size  # Match font size
                    graphic_resized = graphic_image.resize((graphic_size, graphic_size), Image.Resampling.LANCZOS)
                    
                    graphic_x = max(10, int(phone_x) - graphic_size - 5)
                    graphic_y = max(10, int(phone_y) - graphic_size // 4)
                    
                    bg_image.paste(graphic_resized, (graphic_x, graphic_y), graphic_resized)
                except Exception as e:
                    print(f"Error adding graphic: {e}")
            
            return bg_image

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during image processing: {e}")
            return None

    def _setup_ui(self):
        """Sets up the enhanced graphical user interface elements."""
        self.sidebar_frame = ctk.CTkFrame(self.root, width=350)
        self.sidebar_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        self.title_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Enhanced Flyer Generator", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=20)

        self.tab_view = ctk.CTkTabview(self.sidebar_frame)
        self.tab_view.pack(padx=10, pady=10, fill="both", expand=True)

        tabs = [
            ("Files", self._create_files_tab),
            ("Text Style", self._create_text_tab),
            ("Position", self._create_position_tab),
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
            text="Preview", 
            command=self._preview_flyer,
            fg_color="green",
            hover_color="dark green",
            width=140
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            self.button_row1, 
            text="Generate", 
            command=self._generate_flyers,
            fg_color="blue",
            hover_color="dark blue",
            width=140
        ).pack(side="right", padx=5)

        self.whatsapp_buttons_frame = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        self.whatsapp_buttons_frame.pack(fill="x", pady=5)

        self.whatsapp_login_btn = ctk.CTkButton(
            self.whatsapp_buttons_frame,
            text="Connect WhatsApp",
            command=self._login_whatsapp,
            fg_color="green",
            hover_color="dark green",
            width=140
        )
        self.whatsapp_login_btn.pack(side="left", padx=5)

        self.whatsapp_send_btn = ctk.CTkButton(
            self.whatsapp_buttons_frame,
            text="Send Flyers",
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
            text="Use X,Y coordinates in Position tab to move text elements",
            text_color="gray"
        )
        self.instructions_label.grid(row=1, column=0, pady=5)

    def _create_files_tab(self, master):
        """Sets up the controls for file selection."""
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
        """Enhanced text styling controls."""
        font_frame = ctk.CTkFrame(master, fg_color="transparent")
        font_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(font_frame, text="Font", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        if self.font_options:
            font_menu = ctk.CTkComboBox(
                font_frame, 
                values=self.font_options,
                command=lambda choice: [
                    self.selected_font.set(str(self.FONT_FOLDER / choice)),
                    self._update_preview()
                ]
            )
            font_menu.pack(fill="x")
            if self.font_options:
                font_menu.set(self.font_options[0])
        else:
            ctk.CTkLabel(font_frame, text="No fonts found in 'fonts' folder", text_color="red").pack()

        size_frame = ctk.CTkFrame(master, fg_color="transparent")
        size_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(size_frame, text="Font Size", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        size_entry = ctk.CTkEntry(size_frame, textvariable=self.font_size)
        size_entry.pack(fill="x")
        size_entry.bind("<KeyRelease>", lambda e: self._update_preview())

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

        effects_frame = ctk.CTkFrame(master, fg_color="transparent")
        effects_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(effects_frame, text="Text Effects", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=2)
        
        effects_grid = ctk.CTkFrame(effects_frame, fg_color="transparent")
        effects_grid.pack(fill="x")
        
        style_checks = [
            ("Bold", self.text_bold),
            ("Italic", self.text_italic),
            ("Underline", self.text_underline)
        ]
        
        for i, (text, var) in enumerate(style_checks):
            check = ctk.CTkCheckBox(
                effects_grid, 
                text=text, 
                variable=var,
                command=self._update_preview
            )
            check.grid(row=i//2, column=i%2, sticky="w", padx=5, pady=2)

        shadow_frame = ctk.CTkFrame(master, fg_color="transparent")
        shadow_frame.pack(pady=10, fill="x")
        
        shadow_check = ctk.CTkCheckBox(
            shadow_frame,
            text="Text Shadow",
            variable=self.text_shadow,
            command=self._update_preview
        )
        shadow_check.pack(anchor="w", pady=2)
        
        shadow_color_frame = ctk.CTkFrame(shadow_frame, fg_color="transparent")
        shadow_color_frame.pack(fill="x", pady=5)
        
        shadow_entry = ctk.CTkEntry(shadow_color_frame, textvariable=self.shadow_color, width=120)
        shadow_entry.pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            shadow_color_frame,
            text="Shadow Color",
            command=self._choose_shadow_color,
            width=100
        ).pack(side="right")

    def _create_position_tab(self, master):
        """Create coordinate-based positioning controls."""
        ctk.CTkLabel(master, text="Text Positioning", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=10)
        
        name_frame = ctk.CTkFrame(master, fg_color="transparent")
        name_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(name_frame, text="Name Position", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        name_coords_frame = ctk.CTkFrame(name_frame, fg_color="transparent")
        name_coords_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(name_coords_frame, text="X:", width=20).pack(side="left")
        name_x_entry = ctk.CTkEntry(name_coords_frame, textvariable=self.name_x, width=80)
        name_x_entry.pack(side="left", padx=5)
        name_x_entry.bind("<KeyRelease>", lambda e: self._update_coordinates())
        
        ctk.CTkLabel(name_coords_frame, text="Y:", width=20).pack(side="left", padx=(10, 0))
        name_y_entry = ctk.CTkEntry(name_coords_frame, textvariable=self.name_y, width=80)
        name_y_entry.pack(side="left", padx=5)
        name_y_entry.bind("<KeyRelease>", lambda e: self._update_coordinates())
        
        phone_frame = ctk.CTkFrame(master, fg_color="transparent")
        phone_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(phone_frame, text="Phone Position", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        phone_coords_frame = ctk.CTkFrame(phone_frame, fg_color="transparent")
        phone_coords_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(phone_coords_frame, text="X:", width=20).pack(side="left")
        phone_x_entry = ctk.CTkEntry(phone_coords_frame, textvariable=self.phone_x, width=80)
        phone_x_entry.pack(side="left", padx=5)
        phone_x_entry.bind("<KeyRelease>", lambda e: self._update_coordinates())
        
        ctk.CTkLabel(phone_coords_frame, text="Y:", width=20).pack(side="left", padx=(10, 0))
        phone_y_entry = ctk.CTkEntry(phone_coords_frame, textvariable=self.phone_y, width=80)
        phone_y_entry.pack(side="left", padx=5)
        phone_y_entry.bind("<KeyRelease>", lambda e: self._update_coordinates())
        
        quick_pos_frame = ctk.CTkFrame(master, fg_color="transparent")
        quick_pos_frame.pack(pady=15, fill="x")
        
        ctk.CTkLabel(quick_pos_frame, text="Quick Positions", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        positions_grid = ctk.CTkFrame(quick_pos_frame, fg_color="transparent")
        positions_grid.pack(fill="x", pady=5)
        
        quick_positions = [
            ("Top Left", lambda: self._set_quick_position("top_left")),
            ("Top Right", lambda: self._set_quick_position("top_right")),
            ("Bottom Left", lambda: self._set_quick_position("bottom_left")),
            ("Bottom Right", lambda: self._set_quick_position("bottom_right")),
            ("Center", lambda: self._set_quick_position("center"))
        ]
        
        for i, (text, command) in enumerate(quick_positions):
            btn = ctk.CTkButton(
                positions_grid,
                text=text,
                command=command,
                width=90,
                height=30
            )
            btn.grid(row=i//3, column=i%3, padx=2, pady=2, sticky="ew")
        
        for i in range(3):
            positions_grid.grid_columnconfigure(i, weight=1)

    def _set_quick_position(self, position):
        """Set quick positioning presets."""
        if not self.bg_image_path.get():
            return
            
        try:
            bg_image = Image.open(self.bg_image_path.get())
            width, height = bg_image.size
            
            positions = {
                "top_left": {"name": (50, 100), "phone": (50, 140)},
                "top_right": {"name": (width-300, 100), "phone": (width-300, 140)},
                "bottom_left": {"name": (50, height-80), "phone": (50, height-40)},
                "bottom_right": {"name": (width-300, height-80), "phone": (width-300, height-40)},
                "center": {"name": (width//2-100, height//2-20), "phone": (width//2-100, height//2+20)}
            }
            
            pos = positions.get(position, positions["bottom_left"])
            self.name_x.set(str(pos["name"][0]))
            self.name_y.set(str(pos["name"][1]))
            self.phone_x.set(str(pos["phone"][0]))
            self.phone_y.set(str(pos["phone"][1]))
            
            self._update_preview()
            
        except Exception as e:
            print(f"Error setting quick position: {e}")

    def _create_graphics_tab(self, master):
        """Sets up the controls for selecting graphics/icons."""
        ctk.CTkLabel(master, text="Graphics/Icons", font=ctk.CTkFont(size=12, weight="bold")).pack(pady=10)
        
        graphics_files = []
        for ext in ['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp']:
            graphics_files.extend(list(self.GRAPHICS_FOLDER.glob(ext)))
        
        if graphics_files:
            graphics_grid = ctk.CTkScrollableFrame(master, height=300)
            graphics_grid.pack(fill="both", expand=True, padx=5, pady=5)

            for i, graphic in enumerate(graphics_files):
                try:
                    img = ctk.CTkImage(
                        Image.open(graphic).resize((60, 60)), 
                        size=(60, 60)
                    )
                    
                    graphic_frame = ctk.CTkFrame(graphics_grid, fg_color="transparent")
                    graphic_frame.pack(pady=5, fill="x")
                    
                    btn = ctk.CTkButton(
                        graphic_frame, 
                        image=img, 
                        text=graphic.stem[:15], 
                        compound="left",
                        command=lambda g=graphic: self._select_graphic(g),
                        width=200,
                        height=70
                    )
                    btn.pack(fill="x")
                    
                except Exception as e:
                    print(f"Error loading graphic {graphic}: {e}")
        else:
            ctk.CTkLabel(
                master, 
                text="No graphics found.\nAdd images to the 'graphics' folder.",
                text_color="gray"
            ).pack(pady=20)
        
        ctk.CTkButton(
            master,
            text="Clear Graphic",
            command=lambda: [self.selected_graphic.set(""), self._update_preview()],
            fg_color="red",
            hover_color="dark red"
        ).pack(pady=10)

    def _create_whatsapp_tab(self, master):
        """Enhanced WhatsApp automation controls with on/off switches."""
        instructions = ctk.CTkTextbox(master, height=100, wrap="word")
        instructions.pack(pady=10, fill="x")
        instructions.insert("0.0", 
            "WhatsApp Automation:\n\n"
            "1. Connect to WhatsApp Web\n"
            "2. Generate flyers first\n"
            "3. Configure message/caption options below\n"
            "4. Send flyers automatically\n\n"
            "Note: Faster processing, reduced delays for bulk sending."
        )
        instructions.configure(state="disabled")

        message_section = ctk.CTkFrame(master, fg_color="transparent")
        message_section.pack(pady=10, fill="x")
        
        message_header = ctk.CTkFrame(message_section, fg_color="transparent")
        message_header.pack(fill="x")
        
        self.message_switch = ctk.CTkSwitch(
            message_header,
            text="Custom Message",
            variable=self.use_custom_message,
            command=self._toggle_message_controls
        )
        self.message_switch.pack(side="left")
        
        self.message_textbox = ctk.CTkTextbox(message_section, height=60)
        self.message_textbox.pack(fill="x", pady=5)
        self.message_textbox.insert("0.0", self.whatsapp_message.get())
        self.message_textbox.configure(state="disabled")  # Start disabled
        
        def update_message():
            if self.use_custom_message.get():
                self.whatsapp_message.set(self.message_textbox.get("0.0", "end-1c"))
        
        self.message_textbox.bind("<KeyRelease>", lambda e: update_message())

        caption_section = ctk.CTkFrame(master, fg_color="transparent")
        caption_section.pack(pady=10, fill="x")
        
        caption_header = ctk.CTkFrame(caption_section, fg_color="transparent")
        caption_header.pack(fill="x")
        
        self.caption_switch = ctk.CTkSwitch(
            caption_header,
            text="Custom Caption",
            variable=self.use_custom_caption,
            command=self._toggle_caption_controls
        )
        self.caption_switch.pack(side="left")
        
        self.caption_textbox = ctk.CTkTextbox(caption_section, height=50)
        self.caption_textbox.pack(fill="x", pady=5)
        self.caption_textbox.insert("0.0", self.image_caption.get())
        self.caption_textbox.configure(state="disabled")  # Start disabled
        
        def update_caption():
            if self.use_custom_caption.get():
                self.image_caption.set(self.caption_textbox.get("0.0", "end-1c"))
        
        self.caption_textbox.bind("<KeyRelease>", lambda e: update_caption())

        variables_frame = ctk.CTkFrame(master, fg_color="transparent")
        variables_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(
            variables_frame, 
            text="Available Variables:", 
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            variables_frame,
            text="{name} - Contact's name\n{phone} - Contact's phone number",
            text_color="gray",
            justify="left"
        ).pack(anchor="w", pady=2)

        speed_frame = ctk.CTkFrame(master, fg_color="#e6ffe6")
        speed_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(
            speed_frame,
            text="Speed Optimizations Applied:\n• Reduced wait times\n• Parallel processing\n• Faster chat opening\n• Optimized image sending",
            text_color="#006600",
            justify="left"
        ).pack(pady=8)

        warning_frame = ctk.CTkFrame(master, fg_color="#ffebcc")
        warning_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(
            warning_frame,
            text="Important: Use responsibly. Sending too many messages quickly may result in temporary restrictions from WhatsApp.",
            text_color="#cc6600",
            wraplength=280,
            justify="left"
        ).pack(pady=8)

    def _toggle_message_controls(self):
        """Toggle custom message controls on/off."""
        if self.use_custom_message.get():
            self.message_textbox.configure(state="normal")
        else:
            self.message_textbox.configure(state="disabled")

    def _toggle_caption_controls(self):
        """Toggle custom caption controls on/off."""
        if self.use_custom_caption.get():
            self.caption_textbox.configure(state="normal")
        else:
            self.caption_textbox.configure(state="disabled")

def check_dependencies():
    """Check if all required packages are installed."""
    required_packages = {
        'selenium': 'selenium',
        'pandas': 'pandas', 
        'openpyxl': 'openpyxl',
        'customtkinter': 'customtkinter',
        'PIL': 'Pillow',
        'webdriver_manager': 'webdriver-manager',
        'pyperclip': 'pyperclip',
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
    """Initialize and run the enhanced application."""
    if not check_dependencies():
        input("Press Enter to exit...")
        return
    
    print("Starting Enhanced Flyer Generator with WhatsApp Automation...")
    print("Features: Coordinate positioning, text effects, custom messages/captions, faster processing")
    
    app = ModernFlyerGeneratorApp()
    app.root.mainloop()

if __name__ == "__main__":
    main()
