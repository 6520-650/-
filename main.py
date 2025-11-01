from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import os
import sys
import my_zip

class InvoiceDownloader:
    def __init__(self, debug_port=9222, download_path=None):
        self.chrome_options = Options()
        self.chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
        
        # 存储基础下载路径，实际下载路径会根据月份动态创建
        self.base_download_path = download_path
        
        # 初始化时不设置具体下载路径，将在处理每个月份时动态设置
        prefs = {
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        self.chrome_options.add_experimental_option("prefs", prefs)
        
        self.driver = None
        self.wait = None
        self.actions = None
    
    def set_download_path_for_month(self, year, month):
        """为特定月份设置下载路径"""
        if self.base_download_path:
            # 创建月份格式的文件夹名称 (YYYYMM)
            month_folder = f"{year}{month:02d}"
            month_download_path = os.path.join(self.base_download_path, month_folder)
            
            # 创建文件夹（如果不存在）
            os.makedirs(month_download_path, exist_ok=True)
            
            # 动态更新下载路径
            prefs = {
                "download.default_directory": month_download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            
            # 由于Chrome选项在启动后不能直接修改，我们需要通过CDP命令来更新下载路径
            if self.driver:
                try:
                    self.driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                        'behavior': 'allow', 
                        'downloadPath': month_download_path
                    })
                    print(f"📁📁📁📁 下载路径已设置为: {month_download_path}")
                except Exception as e:
                    print(f"⚠️ 无法通过CDP设置下载路径，使用初始路径: {e}")
            
            return month_download_path  # 返回路径供生成报告使用
        
        return None

    def connect_browser(self):
        try:
            self.driver = webdriver.Chrome(options=self.chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            self.actions = ActionChains(self.driver)
            print("✅ 浏览器连接成功")
            return True
        except Exception as e:
            print(f"❌❌ 浏览器连接失败: {e}")
            return False
    
    def navigate_to_page(self, url):
        try:
            print(f"🌐🌐 正在导航到: {url}")
            self.driver.get(url)
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            print("✅ 页面加载成功")
            return True
        except Exception as e:
            print(f"❌❌ 页面导航失败: {e}")
            return False
    
    def click_etc_card(self):
        try:
            # 使用更精确的选择器
            card_xpath = "//a[contains(@href, '广西ETC') or .//dt[contains(text(), '广西ETC')] or contains(text(), '广西ETC')]"
            card_element = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, card_xpath))
            )
            card_element.click()
            print("✅ 已点击广西ETC卡片")
            time.sleep(3)
            return True
        except Exception as e:
            print(f"❌❌ 点击ETC卡失败: {e}")
            return False
    
   
    
    def set_date_js_calendar(self, year, month):
        """使用JavaScript直接调用WdatePicker"""
        try:
            # 构建目标日期字符串
            target_date = f"{year}-{month:02d}"
            
            # 直接调用WdatePicker的setDate方法
            js_code = f"""
            // 创建日期对象
            var targetDate = new Date({year}, {month-1}, 1);
            
            // 查找WdatePicker输入框
            var monthInput = document.getElementById('month');
            if (monthInput) {{
                // 设置值
                monthInput.value = '{year}{month:02d}';
                
                // 触发所有必要的事件
                var events = ['input', 'change', 'blur'];
                events.forEach(function(eventType) {{
                    var event = new Event(eventType, {{ bubbles: true }});
                    monthInput.dispatchEvent(event);
                }});
                
                // 调用可能的回调函数
                if (window.WdatePicker && window.WdatePicker.onpicked) {{
                    window.WdatePicker.onpicked.call(monthInput);
                }}
            }}
            """
            
            self.driver.execute_script(js_code)
            print(f"✅ 已通过JS设置日期: {year}年{month}月")
            time.sleep(2)
            return True
            
        except Exception as e:
            print(f"❌❌ JS设置日期失败: {e}")
            return False
    
    
    
    def set_date(self, year, month):
        """设置日期 - 尝试多种方法"""
        print(f"\n📅📅 开始设置日期: {year}年{month}月")
        
        # 方法2: 使用JavaScript
        if self.set_date_js_calendar(year, month):
            return True
        
        print("❌❌ 所有日期设置方法都失败了")
        return False
    
    def search_invoices(self):
        try:
            # 尝试多种搜索按钮定位方式
            search_selectors = [
                "#titSeach",  # 按ID
                "#seach",     # 备用ID
                "button[type='button']",  # 按钮类型
                ".taiji_search_submit",  # 类名
                "input[value*='搜索']",  # 包含搜索文本
            ]
            
            for selector in search_selectors:
                try:
                    search_button = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    search_button.click()
                    print("🔍🔍 正在搜索发票...")
                    time.sleep(5)
                    return True
                except:
                    continue
            
            # 如果以上都失败，尝试通过JavaScript点击
            js_code = """
            var searchBtn = document.getElementById('titSeach') || 
                          document.getElementById('seach') ||
                          document.querySelector('.taiji_search_submit');
            if (searchBtn) {
                searchBtn.click();
                return true;
            }
            return false;
            """
            
            result = self.driver.execute_script(js_code)
            if result:
                print("🔍🔍 已通过JS点击搜索按钮")
                time.sleep(5)
                return True
            else:
                print("❌❌ 未找到可点击的搜索按钮")
                return False
                
        except Exception as e:
            print(f"❌❌ 搜索失败: {e}")
            return False
    
    def get_invoice_tables(self):
        try:
            # 等待发票表格加载
            time.sleep(3)
            
            # 多种选择器尝试
            table_selectors = [
                "table.table_wdfp",
                "table.table",
                ".table_wdfp",
                "table"
            ]
            
            for selector in table_selectors:
                try:
                    invoice_tables = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if invoice_tables:
                        print(f"📋📋 找到 {len(invoice_tables)} 个发票条目")
                        return invoice_tables
                except:
                    continue
            
            print("❌❌ 未找到发票表格")
            return []
            
        except Exception as e:
            print(f"❌❌ 获取发票表格失败: {e}")
            return []

    def generate_amount_report(self, year, month, download_path, total_amount, invoice_details, success_count, total_count):
        """生成金额统计文件"""
        try:
            # 创建统计文件名
            report_filename = f"{year}{month:02d}_发票统计.txt"
            report_filepath = os.path.join(download_path, report_filename)
            
            with open(report_filepath, 'w', encoding='utf-8') as f:
                f.write(f"发票统计报告 - {year}年{month:02d}月\n")
                f.write("=" * 50 + "\n")
                f.write(f"统计时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"发票总数: {total_count} 张\n")
                f.write(f"下载成功: {success_count} 张\n")
                f.write(f"下载失败: {total_count - success_count} 张\n")
                f.write(f"开票总金额: ￥{total_amount:.2f}\n")
                f.write("\n" + "=" * 50 + "\n")
                f.write("发票明细:\n")
                f.write("-" * 50 + "\n")
                
                for detail in invoice_details:
                    f.write(f"第{detail['index']}张发票: ￥{detail['amount']:.2f} - {detail['status']}\n")
            
            print(f"📄📄📄📄 金额统计文件已生成: {report_filepath}")
            return True
            
        except Exception as e:
            print(f"❌❌❌❌ 生成金额统计文件失败: {e}")
            return False
    
    def download_single_invoice(self, table_element, index):
        try:
            print(f"\n⬇⬇⬇⬇️ 开始处理第 {index} 张发票")
            
            # 提取开票金额
            amount = 0.0
            try:
                # 多种方式查找金额元素
                amount_selectors = [
                    ".//th[contains(., '开票金额')]//span",
                    ".//span[contains(@class, 'inv_deta_list_divc01')]",
                    ".//span[contains(text(), '￥')]",
                    ".//th[contains(., '金额')]//span"
                ]
                
                for selector in amount_selectors:
                    try:
                        amount_element = table_element.find_element(By.XPATH, selector)
                        amount_text = amount_element.text.strip()
                        if '￥' in amount_text:
                            amount_str = amount_text.replace('￥', '').strip()
                            amount = float(amount_str)
                            print(f"💰💰 第 {index} 张发票 - 开票金额: {amount_text}")
                            break
                    except:
                        continue
            except Exception as e:
                print(f"⚠️ 第 {index} 张发票 - 金额提取失败: {e}")
                amount = 0.0
            
            # 多种方式查找下载链接
            link_selectors = [
                ".//a[contains(@href, '/downloadPage/')]",
                ".//a[contains(text(), '下载')]",
                ".//a[contains(@onclick, 'download')]",
                ".//button[contains(text(), '下载')]"
            ]
            
            download_link = None
            for selector in link_selectors:
                try:
                    download_link = table_element.find_element(By.XPATH, selector)
                    break
                except:
                    continue
            
            if not download_link:
                print(f"❌❌❌❌ 第 {index} 张发票 - 未找到下载链接")
                return False, amount
            
            main_window = self.driver.current_window_handle
            
            # 在新标签页中打开下载链接
            self.driver.execute_script("arguments[0].target='_blank';", download_link)
            download_link.click()
            
            print(f"🖱🖱🖱🖱🖱🖱🖱🖱🖱️ 第 {index} 张发票 - 已点击下载链接")
            time.sleep(2)
            
            # 切换到新标签页
            all_windows = self.driver.window_handles
            new_window = [w for w in all_windows if w != main_window]
            
            if new_window:
                self.driver.switch_to.window(new_window[0])
                print(f"✅ 第 {index} 张发票 - 已切换到下载页面")
                
                # 尝试找到打包下载按钮
                download_buttons = [
                    ("ID", "no-invoice"),
                    ("CSS", "input[value*='打包']"),
                    ("CSS", "button[contains(text(), '打包')]"),
                    ("CSS", "input[type='button']"),
                    ("CSS", "button")
                ]
                
                for selector_type, selector_value in download_buttons:
                    try:
                        if selector_type == "ID":
                            download_btn = self.driver.find_element(By.ID, selector_value)
                        else:
                            download_btn = self.driver.find_element(By.CSS_SELECTOR, selector_value)
                        
                        download_btn.click()
                        print(f"📦📦📦📦 第 {index} 张发票 - 已点击下载按钮")
                        break
                    except:
                        continue
                
                time.sleep(2)
                
                # 关闭当前标签页并返回主窗口
                self.driver.close()
                self.driver.switch_to.window(main_window)
                print(f"✅ 第 {index} 张发票下载完成")
                return True, amount
            else:
                print(f"❌❌❌❌ 第 {index} 张发票 - 未打开新标签页")
                return False, amount
                
        except Exception as e:
            print(f"❌❌❌❌ 第 {index} 张发票下载失败: {e}")
            # 确保返回主窗口
            try:
                self.driver.switch_to.window(self.driver.window_handles[0])
            except:
                pass
            return False, 0.0

    def process_single_month(self, year, month):
        """处理单个月份的发票下载"""
        print(f"\n{'='*60}")
        print(f"📅📅📅📅 开始处理 {year}年{month:02d}月 的发票")
        print(f"{'='*60}")
        
        # 首先设置该月份的下载路径
        month_download_path = self.set_download_path_for_month(year, month)
        
        # 设置日期
        if not self.set_date(year, month):
            print(f"⚠️ {year}年{month:02d}月 - 日期设置可能失败，但继续尝试搜索...")
        
        if not self.search_invoices():
            print(f"❌❌❌❌ {year}年{month:02d}月 - 搜索失败")
            return False
        
        invoice_tables = self.get_invoice_tables()
        if not invoice_tables:
            print(f"❌❌❌❌ {year}年{month:02d}月 - 未找到可下载的发票")
            return False
        
        print(f"🎯🎯🎯🎯 {year}年{month:02d}月 - 开始批量下载，共 {len(invoice_tables)} 张发票")
        
        success_count = 0
        total_amount = 0.0  # 总金额统计
        invoice_details = []  # 发票明细
        
        for i in range(len(invoice_tables)):
            # 重新获取表格元素，避免StaleElementReferenceException
            current_tables = self.get_invoice_tables()
            if i < len(current_tables):
                success, amount = self.download_single_invoice(current_tables[i], i+1)
                if success:
                    success_count += 1
                if amount > 0:
                    total_amount += amount
                    invoice_details.append({
                        'index': i+1,
                        'amount': amount,
                        'status': '成功' if success else '失败'
                    })
            
            if i < len(invoice_tables) - 1:
                print("⏳⏳⏳⏳⏳⏳⏳⏳⏳ 等待3秒后处理下一张发票...")
                time.sleep(2)
        
        # 生成金额统计文件
        self.generate_amount_report(year, month, month_download_path, total_amount, invoice_details, success_count, len(invoice_tables))
        
        print(f"\n📊📊📊📊 {year}年{month:02d}月 - 下载完成!")
        print(f"   成功: {success_count} 张")
        print(f"   失败: {len(invoice_tables) - success_count} 张")
        print(f"   总金额: ￥{total_amount:.2f}")
        
        return success_count > 0

    def batch_download(self, target_url, year, month):
        """单个月份下载的兼容方法"""
        if not self.connect_browser():
            return False
        
        try:
            if not self.navigate_to_page(target_url):
                return False
            
            if not self.click_etc_card():
                return False
            
            return self.process_single_month(year, month)
            
        except Exception as e:
            print(f"❌❌ 批量下载过程出错: {e}")
            return False

    def batch_download_multiple_months(self, target_url, month_list):
        """批量下载多个月份的发票"""
        if not self.connect_browser():
            return False
        
        try:
            overall_success = True
            total_months = len(month_list)
            
            for idx, (year, month) in enumerate(month_list, 1):
                print(f"\n📊📊 进度: 第 {idx}/{total_months} 个月份")
                
                # 每次处理新月份时都重新导航到初始页面
                if not self.navigate_to_page(target_url):
                    print(f"❌❌ {year}年{month:02d}月 - 页面导航失败")
                    overall_success = False
                    continue
                
                if not self.click_etc_card():
                    print(f"❌❌ {year}年{month:02d}月 - ETC卡片点击失败")
                    overall_success = False
                    continue
                
                month_success = self.process_single_month(year, month)
                if not month_success:
                    overall_success = False
                
                # 如果不是最后一个月份，等待一段时间再处理下一个月份
                if idx < total_months:
                    wait_time = 3  # 等待3秒再处理下一个月
                    print(f"⏳⏳⏳ 等待{wait_time}秒后处理下一个月份...")
                    time.sleep(wait_time)
            
            return overall_success
            
        except Exception as e:
            print(f"❌❌ 批量下载过程出错: {e}")
            return False
    
    def close(self):
        if self.driver:
            self.driver.quit()
            print("🔚🔚 浏览器已关闭")

def parse_month_input(month_str):
    """解析月份输入，支持多种格式"""
    month_str = month_str.strip()
    
    # 格式1: YYYYMM (如 202410)
    if len(month_str) == 6 and month_str.isdigit():
        year = int(month_str[:4])
        month = int(month_str[4:6])
        return year, month
    
    # 格式2: YYYY-MM (如 2024-10)
    elif '-' in month_str:
        parts = month_str.split('-')
        if len(parts) == 2 and all(part.isdigit() for part in parts):
            year = int(parts[0])
            month = int(parts[1])
            return year, month
    
    # 格式3: YYYY年M月 (如 2024年10月)
    elif '年' in month_str and '月' in month_str:
        year_str = month_str.split('年')[0]
        month_str_clean = month_str.split('年')[1].replace('月', '')
        if year_str.isdigit() and month_str_clean.isdigit():
            year = int(year_str)
            month = int(month_str_clean)
            return year, month
    
    raise ValueError(f"无法解析的月份格式: {month_str}")

def get_month_range(start_year, start_month, end_year, end_month):
    """生成月份范围列表"""
    months = []
    current_year, current_month = start_year, start_month
    
    while (current_year < end_year) or (current_year == end_year and current_month <= end_month):
        months.append((current_year, current_month))
        
        current_month += 1
        if current_month > 12:
            current_month = 1
            current_year += 1
    
    return months

def main():
    DEBUG_PORT = 9222 #默认端口号根据调试端口填写
    TARGET_URL = "https://pss.txffp.com/pss/app/login/invoice/query/card/PERSONAL"
    DOWNLOAD_PATH = os.path.join(os.getcwd(), "invoice_downloads")
    
    print("批量下载")
    print("=" * 50)
    print("请选择下载模式:")
    print("1. 单个月份下载")
    print("2. 多个月份批量下载")
    print("3. 连续月份范围下载")
    print("4. ptf解压")
    
    mode_choice = input("请选择模式 (1/2/3/4): ").strip()
    
    os.makedirs(DOWNLOAD_PATH, exist_ok=True)
    
    downloader = InvoiceDownloader(DEBUG_PORT, DOWNLOAD_PATH)
    
    try:
        if mode_choice == "1":
            # 单个月份下载模式
            print("\n📅📅 单个月份下载模式")
            print("请输入要下载发票的年月信息 (格式: YYYYMM)")
            print("例如: 2025年10月请输入 202510")
            
            while True:
                try:
                    date_input = input("日期 (YYYYMM): ").strip()
                    year, month = parse_month_input(date_input)
                    if 2020 <= year <= 2030 and 1 <= month <= 12:
                        break
                    else:
                        print("❌❌ 日期范围无效，请输入2020-2030年的有效月份")
                except Exception as e:
                    print(f"❌❌ 输入错误: {e}，请使用YYYYMM格式重新输入")
            
            print(f"\n🚀🚀 开始下载 {year}年{month:02d}月 的发票")
            downloader.batch_download(TARGET_URL, year, month)
            
            # 下载完成后直接调用解压
            print("\n📦📦 下载任务完成，开始解压文件...")
            my_zip.main("invoice_downloads")
            print("✅ 解压完成!")
            
        elif mode_choice == "2":
            # 多个月份批量下载模式
            print("\n📅📅 多个月份批量下载模式")
            print("请输入要下载的多个月份，用逗号或空格分隔")
            print("格式示例: 202410,202411,202412 或 202410 202411 202412")
            print("或: 2024-10,2024-11,2024-12")
            
            month_input = input("月份列表: ").strip()
            
            # 处理分隔符
            separators = [',', '，', ' ', ';']
            for sep in separators:
                if sep in month_input:
                    month_list_str = [x.strip() for x in month_input.split(sep) if x.strip()]
                    break
            else:
                month_list_str = [month_input]
            
            month_list = []
            for month_str in month_list_str:
                try:
                    year, month = parse_month_input(month_str)
                    if 2020 <= year <= 2030 and 1 <= month <= 12:
                        month_list.append((year, month))
                    else:
                        print(f"❌❌ 跳过无效日期: {month_str}")
                except Exception as e:
                    print(f"❌❌ 跳过无法解析的日期: {month_str} - {e}")
            
            if not month_list:
                print("❌❌ 没有有效的月份输入，程序退出")
                return
            
            print(f"\n🎯🎯 将要下载以下 {len(month_list)} 个月的发票:")
            for year, month in month_list:
                print(f"  - {year}年{month:02d}月")
            
            confirm = input("\n确认开始下载? (y/N): ").strip().lower()
            if confirm != 'y':
                print("下载已取消")
                return
            
            downloader.batch_download_multiple_months(TARGET_URL, month_list)
            
            # 下载完成后直接调用解压
            print("\n📦📦 下载任务完成，开始解压文件...")
            my_zip.main("invoice_downloads")
            print("✅ 解压完成!")
            
        elif mode_choice == "3":
            # 连续月份范围下载模式
            print("\n📅📅 连续月份范围下载模式")
            print("请输入起始月份和结束月份")
            
            start_input = input("起始月份 (YYYYMM): ").strip()
            end_input = input("结束月份 (YYYYMM): ").strip()
            
            try:
                start_year, start_month = parse_month_input(start_input)
                end_year, end_month = parse_month_input(end_input)
                
                if start_year > end_year or (start_year == end_year and start_month > end_month):
                    print("❌❌ 起始月份不能晚于结束月份")
                    return
                
                month_list = get_month_range(start_year, start_month, end_year, end_month)
                
                print(f"\n🎯🎯 将要下载从 {start_year}年{start_month:02d}月 到 {end_year}年{end_month:02d}月 的发票")
                print(f"共计 {len(month_list)} 个月份")
                
                confirm = input("\n确认开始下载? (y/N): ").strip().lower()
                if confirm != 'y':
                    print("下载已取消")
                    return
                
                downloader.batch_download_multiple_months(TARGET_URL, month_list)
                
                # 下载完成后直接调用解压
                print("\n📦📦 下载任务完成，开始解压文件...")
                my_zip.main("invoice_downloads")
                print("✅ 解压完成!")
                
            except Exception as e:
                print(f"❌❌ 日期解析错误: {e}")
                return
        
        elif mode_choice == "4":
            my_zip.main("invoice_downloads")
            return        
        else:
            print("❌❌ 无效的选择，请选择1、2或3/4")
            return
            
    except Exception as e:
        print(f"\n💥💥 脚本执行异常: {e}")
    finally:
        downloader.close()

if __name__ == "__main__":
    print("通用文件处理工具示例")
    print(" 仅用于学习和研究目的")
    main()
    input("\n\n程序执行完毕，按回车键退出...")