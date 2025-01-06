import os
import re
import time
import requests
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

file_path = r'C:\Users\HemantTank\Desktop\back_up\Data Upload\upload\6_12_output_www.ajmadison.com_data_all 2.xlsx'

if not os.path.exists("ajmadison_com"):
    os.makedirs("ajmadison_com")
output_directory = os.path.join(os.getcwd(), 'ajmadison_com')

driver = webdriver.Chrome()

driver.get('https://account-management.moveeasy.com/account-management/missing-appliance-data-list/')

email_input = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, '//input[@name="email"]'))
)
email_input.send_keys('#######################')
 
password_input = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, '//input[@name="password"]'))
)
password_input.send_keys('########################')
 
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@id="id_login"]'))
)
login_button.click()
time.sleep(10)
driver.get("https://account-management.moveeasy.com/account-management/missing-manual-view/")
time.sleep(5)
 
 
df_pdf_data = pd.read_excel(file_path, sheet_name='Sheet1')
 
for index, row in df_pdf_data[:].iterrows():
    model_no = re.sub(r'\s+', ' ', str(row['model_name'])).strip()
    name = re.sub(r'\s+', ' ', str(row['name'])).strip()
    brand = re.sub(r'\s+', ' ', str(row['brand'])).strip()
    if brand:
        if "VIKING" in brand.upper():
            brand = "VIKING"
    category = re.sub(r'\s+', ' ', str(row['category'])).strip()
    if category:
        if "BLOWERS" in category.upper():
            category = "Others"
    # category = "Ranges & Ovens"
    image_url = re.sub(r'\s+', ' ', str(row['image_url'])).strip()
    user_manual = re.sub(r'\s+', ' ', str(row['user_manual'])).strip()
    Energy_Guide = re.sub(r'\s+', ' ', str(row['Energy_Guide'])).strip()
    Warranty_Guide = re.sub(r'\s+', ' ', str(row['Warranty_Guide'])).strip()
    print(model_no)
    print(image_url)
    print(brand)
    print(name)
    print(category)
    print(user_manual)
    print(Warranty_Guide)
    print(Energy_Guide)
    if name != "nan":
        print("AAAAAAAAAAAAA")
# //tbody/tr/td[text()={model_no}]/parent::tr/td/a
# model_no = "MDB4409PAWO"
# serial_no = "F21004806"
 
    try:
        model_search_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//input[@type="search"]'))
        )
        model_search_input.clear()
        model_search_input.send_keys(model_no)
        model_search_input.send_keys(Keys.RETURN)
        time.sleep(3)
        clicked_on_model = False
        try:
            button_all = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//select[@id="id_appliance_user_manual_status"]/option[contains(text(), "All")]'))
            )
    
            button_all.click()
        except:
            pass
        pagination_count = 1
        time.sleep(2)
        try:
            previous_page_text = driver.find_element(By.XPATH, "//li[@class='paginate_button page-item next']/preceding-sibling::li[1]/a").text
            print(previous_page_text)
            pagination_count = int(previous_page_text)
        except:
            pagination_count = 1
        
        for i in range(0,pagination_count):
            try:
                button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f'//tbody/tr/td[text()="{model_no}"]/parent::tr/td/a'))
                )
        
                button.click()
                time.sleep(6)
                clicked_on_model = True
                break
            except:
                try:
                    next_page_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, f'//li[@class="paginate_button page-item next"]/a[contains(text(), "Next")]'))
                    )
            
                    next_page_button.click()
                except:
                    pass
        if clicked_on_model == False:
            pagination_count_has_manual = 1
            try:
                button_manual = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f'//select[@id="id_appliance_user_manual_status"]/option[@value="Has User Manual"]'))
                )
        
                button_manual.click()
            except:
                pass
            time.sleep(2)
            try:
                previous_page_text = driver.find_element(By.XPATH, "//li[@class='paginate_button page-item next']/preceding-sibling::li[1]/a").text
                print(previous_page_text)
                pagination_count_has_manual = int(previous_page_text)
            except:
                pagination_count_has_manual = 1
            
            for j in range(0,pagination_count_has_manual):
                try:
                    button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, f'//tbody/tr/td[text()="{model_no}"]/parent::tr/td/a'))
                    )
            
                    button.click()
                    time.sleep(6)
                    clicked_on_model = True
                    break
                except:
                    next_page_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, f'//li[@class="paginate_button page-item next"]/a[contains(text(), "Next")]'))
                    )
            
                    next_page_button.click()
        windows = driver.window_handles
        print("Window handles:", windows)
        driver.switch_to.window(windows[1])
        time.sleep(5)
        # if model_no != "nan":
        #     model_condition = ""
        #     try:
        #         model_element = WebDriverWait(driver, 10).until(
        #             EC.presence_of_element_located((By.XPATH, f'//input[@name="model_no"]'))
        #         )
        #         model_value = model_element.get_attribute('value')
        #         if model_value == model_no:
        #             model_condition = True
        #         else:
        #             model_condition = False
 
        #     except Exception as e:
        #         print(f"Error: {e}")
        #         model_condition = False
        #     print(model_condition)
        #     if model_condition == False:
        #         Model_field = WebDriverWait(driver, 10).until(
        #             EC.visibility_of_element_located((By.XPATH, '//input[@name="model_no"]'))
        #         )
        #         Model_field.send_keys(model_no)
        #         time.sleep(2)
        if brand != "nan":
            brand_condition = ""
            try:
                select_brand_check = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, f'//select[@name="brand"]/following-sibling::span/span/span/ul/li'))
                )
                if len(select_brand_check) == 2:
                    brand_condition = True
                    print("brand_condition", brand_condition)
                else:
                    select_brand = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH,f'//li[contains(text(), "{brand}")]'))
                    )
                    actions = ActionChains(driver)
                    actions.move_to_element(select_brand).click().perform()
                    brand_condition = False
 
            except Exception as e:
                print(f"Error: {e}")
                brand_condition = False
        # if category != "nan":
        #     category_condition = ""
        #     try:
        #         select_category_check = WebDriverWait(driver, 10).until(
        #             EC.presence_of_all_elements_located((By.XPATH, f'//select[@name="category"]/following-sibling::span/span/span/ul/li'))
        #         )
        #         if len(select_category_check) == 2:
        #             category_condition = True
        #             print("category_condition", category_condition)
        #         else:
        #             select_category = WebDriverWait(driver, 10).until(
        #                 EC.element_to_be_clickable((By.XPATH,f'//li[contains(text(), "{category}")]'))
        #             )
        #             actions = ActionChains(driver)
        #             actions.move_to_element(select_category).click().perform()
        #             category_condition = False
 
        #     except Exception as e:
        #         print(f"Error: {e}")
        #         category_condition = False
        if name != "nan":
            name_condition = ""
            try:
                name_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'//input[@name="name"]'))
                )
                name_value = name_element.get_attribute('value')
                print(name_value)
                if name_value:
                    name_condition = True
                else:
                    name_condition = False
 
            except Exception as e:
                try:
                    name_field = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//input[@name="name"]'))
                    )
                    name_field.send_keys(name)
                    time.sleep(2)
                    print(f"Error: {e}")
                    name_condition = False
                except:
                    name_condition = False
        if user_manual != "nan":
            manual_condition = ""
            try:
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@name="user_manual"]'))
                )
                link_manual = driver.find_element(By.XPATH, '//input[@name="user_manual"]/preceding-sibling::a[1]')
                href_value = link_manual.get_attribute('href')
               
                if href_value:
                    print(href_value)
                    manual_condition = True
                else:
                    manual_condition = False
                    print("No href attribute found.")
               
            except Exception as e:
                print("Error:", e)
            if manual_condition == False:
                try:
                    if user_manual != "":
                        pdf_response = requests.get(user_manual)
 
                        if pdf_response.status_code == 200:
                            model_name_file = re.sub(r'[<>:"/\\|?*]', '_', model_no)
                            model_dir_path = os.path.join(output_directory, model_name_file.replace("/", "_"))
                            if not os.path.exists(model_dir_path):
                                os.makedirs(model_dir_path)
                            pdf_filename = os.path.join(model_dir_path, f"{index}_{model_name_file.replace("/", "_")}user_manual_EN.pdf")
                           
                            with open(pdf_filename, 'wb') as f:
                                f.write(pdf_response.content)
                            user_manual_element = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.XPATH, '//input[@name="user_manual"]'))
                            )
 
                            # Now send the string to the input field
                            user_manual_element.send_keys(pdf_filename)
                except:
                    user_manual = ""
        time.sleep(2)
        if image_url != "nan":
 
            image_condition = ""
            try:
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@name="appliance_image"]'))
                )
                link_image = driver.find_element(By.XPATH, '//input[@name="appliance_image"]/preceding-sibling::a[1]')
                href_value = link_image.get_attribute('href')
               
                if href_value:
                    print(href_value)
                    image_condition = True
                else:
                    image_condition = False
                    print("No href attribute found.")
               
            except Exception as e:
                print("Error:", e)
            if image_condition == False:
                try:
                    if image_url != "":
                        response = requests.get(image_url)
 
                        if response.status_code == 200:
                            model_name_file = re.sub(r'[<>:"/\\|?*]', '_', model_no)
                            model_dir_path = os.path.join(output_directory, model_name_file.replace("/", "_"))
                            if not os.path.exists(model_dir_path):
                                os.makedirs(model_dir_path)
                            image_filename = os.path.join(model_dir_path, f"{index}_{model_name_file.replace("/", "_")}_image.jpg")
                           
                            with open(image_filename, 'wb') as f:
                                f.write(response.content)
                            print(image_url, 'image_url')
                            image_input = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.XPATH, '//input[@name="appliance_image"]'))
                            )
 
                            image_input.send_keys(image_filename)
                except:
                    image_url = ""
 
        time.sleep(2)
        if Warranty_Guide != "nan":
            warranty_condition = ""
            try:
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@name="warranty_doc"]'))
                )
                link_Warranty = driver.find_element(By.XPATH, '//input[@name="warranty_doc"]/preceding-sibling::a[1]')
                href_value = link_Warranty.get_attribute('href')
               
                if href_value:
                    print(href_value)
                    warranty_condition = True
                else:
                    warranty_condition = False
                    print("No href attribute found.")
               
            except Exception as e:
                print("Error:", e)
            if warranty_condition == False:
                try:
                    if  Warranty_Guide != "":
                        Warranty_pdf = requests.get(Warranty_Guide)
 
                        if Warranty_pdf.status_code == 200:
                            model_name_file = re.sub(r'[<>:"/\\|?*]', '_', model_no)
                            model_dir_path = os.path.join(output_directory, model_name_file.replace("/", "_"))
                            if not os.path.exists(model_dir_path):
                                os.makedirs(model_dir_path)
                            pdf_filename = os.path.join(model_dir_path, f"{index}_{model_name_file.replace("/", "_")}Warranty_Guide_EN.pdf")
                           
                            with open(pdf_filename, 'wb') as f:
                                f.write(Warranty_pdf.content)
                            Warranty_Guide_element = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.XPATH, '//input[@name="warranty_doc"]'))
                            )
 
                            # Now send the string to the input field
                            Warranty_Guide_element.send_keys(pdf_filename)
                except:
                    print("EEEEEEEEEEEEEEEEEEEEEEEEE")
                    Warranty_Guide = ""
            time.sleep(2)
        if Energy_Guide != "nan":
            Energy_condition = ""
            try:
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@name="energy_guide"]'))
                )
                link_Energy = driver.find_element(By.XPATH, '//input[@name="energy_guide"]/preceding-sibling::a[1]')
                href_value = link_Energy.get_attribute('href')
               
                if href_value:
                    print(href_value)
                    Energy_condition = True
                else:
                    Energy_condition = False
                    print("No href attribute found.")
               
            except Exception as e:
                print("Error:", e)
            if Energy_condition == False:
                try:
                    if  Energy_Guide != "":
                        Energy_pdf = requests.get(Energy_Guide)
 
                        if Energy_pdf.status_code == 200:
                            model_name_file = re.sub(r'[<>:"/\\|?*]', '_', model_no)
                            model_dir_path = os.path.join(output_directory, model_name_file.replace("/", "_"))
                            if not os.path.exists(model_dir_path):
                                os.makedirs(model_dir_path)
                            pdf_filename = os.path.join(model_dir_path, f"{index}_{model_name_file.replace("/", "_")}Energy_Guide_EN.pdf")
                           
                            with open(pdf_filename, 'wb') as f:
                                f.write(Energy_pdf.content)
                            Energy_Guide_element = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.XPATH, '//input[@name="energy_guide"]'))
                            )
 
                            Energy_Guide_element.send_keys(pdf_filename)
                except:
                    Energy_Guide = ""
 
        update_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//button[contains(text(), "Data Update")]'))
        )
 
        update_button.click()
        time.sleep(10)
       
        driver.close()
 
        # Optionally, switch back to the first window if needed
        driver.switch_to.window(windows[0])
        time.sleep(5)
        if image_url or user_manual or Energy_Guide or Warranty_Guide:
            button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//tbody/tr/td[text()="{model_no}"]/parent::tr/td/select/option[contains(text(), "Reviewed ")]'))
            )
            button.click()
            print("yess review")
        time.sleep(10)
    except Exception as e:
        try:
            print(f"error in full body Upload {e}")
            driver.switch_to.window(windows[0])
        except Exception as e:
            print(f"error in windows change else block {e}")