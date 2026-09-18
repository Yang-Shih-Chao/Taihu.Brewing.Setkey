import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from playwright.sync_api import sync_playwright
import threading
import time
import datetime
import os

def start_automation(file_path, mode="全部新增"):
    if mode not in ["新增子桶", "新增套餐", "全部新增"]:
        messagebox.showinfo("功能開發中", f"目前「{mode}」的功能尚未實作，現階段所有功能均屬於「新增子桶」。")
        return

    # ---------------------------------------------------------
    # 1. 讀取資料 (支援 Excel 與 CSV)
    # ---------------------------------------------------------
    try:
        if file_path.lower().endswith('.csv'):
            # 【重要紀錄】: 在台灣 Windows 環境下，Excel 匯出的 CSV 預設編碼通常是 Big5 (cp950)。
            # 如果直接用預設的 UTF-8 讀取會發生 `UnicodeDecodeError`。
            # 因此這裡實作了自動降級：先嘗試標準的 utf-8-sig，失敗就自動切換成 big5。
            try:
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='big5')
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取檔案失敗:\n{e}")
        return

    df.columns = df.columns.str.strip()

    # 檢查必填欄位是否存在
    required_cols = ["商品編號", "商品名稱", "單價"]
    for col in required_cols:
        if col not in df.columns:
            messagebox.showerror("錯誤", f"檔案內找不到必填欄位: 「{col}」\n請檢查標題列是否正確。")
            return

    messagebox.showinfo("啟動提示", "即將開啟瀏覽器並執行自動化流程。\n(包含自動登入與自動填表)")

    report_lines = [] # 用來記錄每一筆執行結果的清單

    # ---------------------------------------------------------
    # 2. 啟動 Playwright 瀏覽器自動化
    # ---------------------------------------------------------
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # 進入登入頁面
        page.goto("https://hq.caterlord.com/")
        page.wait_for_load_state("networkidle")
        
        # 【自動登入邏輯】
        print("正在自動登入...")
        try:
            page.locator('input[type="text"], input[name*="user" i]').first.fill("demotaihu")
            page.locator('input[type="password"]').first.fill("demotaihu")
            page.locator('button[type="submit"], input[type="submit"], a:has-text("Login"), a:has-text("登入"), button:has-text("Login"), button:has-text("登入")').first.click()
            page.wait_for_load_state("networkidle")
        except Exception as e:
            messagebox.showinfo("自動登入失敗", "找不到登入欄位，請手動登入！\n登入完成後請點擊「確定」繼續執行。")

        # 切換到指定帳號
        print("正在切換帳號...")
        page.goto("https://hq.caterlord.com/CommonTools/SwitchAccount?accountId=12643")
        page.wait_for_load_state("networkidle")
        if mode == "新增套餐":
            print("開始執行「新增套餐」模式...")
            page.goto("https://hq.caterlord.com/Set/SetGroupIndex/")
            page.wait_for_load_state("networkidle")
            
            for index, row in df.iterrows():
                # 使用者的特別欄位對應規則
                group_name = str(row.get("商品名稱", "")).strip()
                group_name_alt = str(row.get("商品編號", "")).strip()
                
                # 數量直接寫死為 1
                min_count = "1"
                max_count = "1"
                
                if not group_name or group_name.lower() in ['nan', 'none']:
                    continue
                
                if not group_name_alt or group_name_alt.lower() in ['nan', 'none']:
                    continue

                # 判斷邏輯：若是 SD、SR、SF 結尾則跳過，只有 S 結尾才執行
                if group_name_alt.endswith(('SD', 'SR', 'SF')):
                    print(f"␐跳過␑商品編號為 SD/SR/SF 結尾: {group_name_alt}")
                    continue
                if not group_name_alt.endswith('S'):
                    print(f"␐跳過␑商品編號非 S 結尾: {group_name_alt}")
                    continue
                    
                print(f"正在處理套餐組合: {group_name}")
                
                try:
                    page.locator("a:has-text('新增套餐項目組合'), .k-grid-add").first.click()
                    page.wait_for_timeout(1000)
                    
                    page.locator("#GroupBatchName").fill(group_name)
                    if group_name_alt and group_name_alt.lower() not in ['nan', 'none']:
                        page.locator("#GroupBatchNameAlt").fill(group_name_alt)
                    
                    if min_count and min_count.lower() not in ['nan', 'none']:
                        page.evaluate(f"() => {{ var tb = $('#MinModifierSelectCount').data('kendoNumericTextBox'); if(tb) {{ tb.value({min_count}); tb.trigger('change'); }} }}")
                    if max_count and max_count.lower() not in ['nan', 'none']:
                        page.evaluate(f"() => {{ var tb = $('#MaxModifierSelectCount').data('kendoNumericTextBox'); if(tb) {{ tb.value({max_count}); tb.trigger('change'); }} }}")
                        
                    page.locator("a.k-grid-update").first.click()
                    page.wait_for_timeout(1000)
                    page.wait_for_load_state("networkidle")
                    
                    msg = f"✅ 【成功】套餐組合 (名稱: {group_name})"
                    print(msg)
                    report_lines.append(msg)
                    
                    # --- 回到 SetGroupIndex 頁面，用篩選功能找到該筆 ---
                    print(f"正在回到套餐組合列表，用篩選功能尋找商品編號: {group_name_alt}")
                    page.goto("https://hq.caterlord.com/Set/SetGroupIndex/")
                    page.wait_for_load_state("networkidle")
                    page.reload()
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(1000)
                    
                    # 1. 點擊「項目組合名稱 (第二語言)」欄位的篩選圖示（漏斗 icon）
                    page.locator("xpath=//th[contains(., '項目組合名稱 (第二語言)')]//a[contains(@class, 'k-grid-filter')]").click()
                    page.wait_for_timeout(500)
                    
                    # 2. 在篩選彈出框中輸入商品編號
                    page.locator("xpath=//input[@title='值']").first.fill(group_name_alt)
                    page.wait_for_timeout(300)
                    
                    # 3. 按下「過濾」按鈕
                    page.locator("xpath=//button[@title='過濾']").click()
                    page.wait_for_timeout(1000)
                    page.wait_for_load_state("networkidle")
                    
                    # 4. 篩選後精準定位該列，點擊展開箭頭
                    target_row = page.locator(f"xpath=//td[@role='gridcell' and normalize-space(text())='{group_name_alt}']/parent::tr")
                    target_row.first.locator("a.k-icon.k-i-expand").click()
                    page.wait_for_timeout(1500)
                    
                    # 5. 展開後，點擊「加入項目」按鈕
                    page.locator("xpath=//a[contains(@class, 'k-grid-AddItem')]").first.click()
                    page.wait_for_timeout(2000)
                    
                    print("已點擊「加入項目」，等待彈出視窗...")
                    
                    # 6. 在彈出的「選擇套餐項目」視窗中，點擊「項目編碼」欄位的篩選圖示
                    # 彈出視窗的 ID 通常是 setGroupItemSelectorWindow
                                        # 6. 使用 JavaScript 直接強制點擊「項目編碼」的篩選漏斗，無視任何動畫或隱藏層阻擋
                    js_click_filter = '''
                    () => {
                        var windows = document.querySelectorAll('.k-window');
                        var visibleWindow = null;
                        for(var i=0; i<windows.length; i++) {
                            var w = windows[i];
                            if(w.offsetWidth > 0 && w.offsetHeight > 0 && w.style.display !== 'none') {
                                visibleWindow = w;
                                break;
                            }
                        }
                        if(visibleWindow) {
                            var ths = visibleWindow.querySelectorAll('th');
                            for(var j=0; j<ths.length; j++) {
                                if(ths[j].innerText.includes('項目編碼')) {
                                    var filterLink = ths[j].querySelector('a.k-grid-filter');
                                    if(filterLink) {
                                        filterLink.click();
                                        return true;
                                    }
                                }
                            }
                        }
                        return false;
                    }
                    '''
                    page.evaluate(js_click_filter)
                    page.wait_for_timeout(1000)
                    
                    # 7. 在篩選選單輸入框中填入商品編號 (直接找畫面上正在顯示的、且可以用來輸入文字的篩選框)
                    # 使用 Playwright 支援的 :visible 偽類來過濾可見元素，避免 filter() 語法錯誤
                    visible_input = page.locator("input[title='值']:visible").last
                    visible_input.fill(group_name_alt)
                    page.wait_for_timeout(500)
                    
                    # 8. 點擊「過濾」按鈕
                    visible_btn = page.locator("button[title='過濾']:visible, button:has-text('過濾'):visible").last
                    visible_btn.click()
                    page.wait_for_timeout(1500)
                    page.wait_for_load_state("networkidle")
                    
                    print(f"✅ 彈出視窗已成功篩選項目編碼: {group_name_alt}")
                    
                    print(f"執行打勾邏輯... (商品編號: {group_name_alt})")
                    js_check_logic = f"""
                    () => {{
                        var groupNameAlt = "{group_name_alt}".toUpperCase();
                        var isGroupStartsS = groupNameAlt.startsWith('S');
                        
                        var windows = document.querySelectorAll('.k-window');
                        var visibleWindow = null;
                        for(var i=0; i<windows.length; i++) {{
                            var w = windows[i];
                            if(w.offsetWidth > 0 && w.offsetHeight > 0 && w.style.display !== 'none') {{
                                visibleWindow = w;
                                break;
                            }}
                        }}
                        
                        if(visibleWindow) {{
                            var rows = visibleWindow.querySelectorAll('.k-grid-content tbody tr');
                            for(var j=0; j<rows.length; j++) {{
                                var row = rows[j];
                                var cells = row.querySelectorAll('td');
                                if(cells.length === 0) continue;
                                
                                var itemCode = cells[0].innerText.trim().toUpperCase();
                                var isItemStartsS = itemCode.startsWith('S');
                                var isItemEndsWithTarget = itemCode.endsWith('SD') || itemCode.endsWith('SR') || itemCode.endsWith('SF');
                                
                                var shouldCheck = false;
                                if (isGroupStartsS) {{
                                    if (isItemEndsWithTarget && isItemStartsS) shouldCheck = true;
                                }} else {{
                                    if (isItemEndsWithTarget && !isItemStartsS) shouldCheck = true;
                                }}
                                
                                if (shouldCheck) {{
                                    var checkbox = row.querySelector("input[type='checkbox']");
                                    var label = row.querySelector("label.chkbx-label");
                                    if (checkbox && !checkbox.checked) {{
                                        if (label) label.click();
                                        else checkbox.click();
                                    }}
                                }}
                            }}
                        }}
                    }}
                    """
                    page.evaluate(js_check_logic)
                    page.wait_for_timeout(1000)
                    
                    # 9. 點擊彈出視窗內的「儲存」按鈕
                    print("點擊彈出視窗內的「儲存」按鈕...")
                    page.locator(".k-window:visible button:has-text('儲存'), .k-window:visible a:has-text('儲存')").first.click()
                    page.wait_for_timeout(2000)
                    
                    # 10. 點擊展開列內的「儲存」按鈕以完成加入項目
                    print("點擊展開列內的「儲存」按鈕...")
                    page.locator(".k-detail-row:visible a.k-grid-save-changes, .k-detail-row:visible a:has-text('儲存')").first.click()
                    page.wait_for_timeout(2000)
                    page.wait_for_load_state("networkidle")
                    
                    msg = f"✅ 【成功】套餐組合 (名稱: {group_name}) 已完成項目勾選與儲存"
                    print(msg)
                    report_lines.append(msg)
                except Exception as e:
                    msg = f"❌ 【失敗】套餐組合 (名稱: {group_name}) - 錯誤: {str(e).splitlines()[0]}"
                    print(msg)
                    report_lines.append(msg)
                    
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            report_filename = f"匯入報告_{timestamp}.txt"
            script_dir = os.path.dirname(os.path.abspath(__file__))
            report_path = os.path.join(script_dir, report_filename)
            with open(report_path, "w", encoding="utf-8") as f:
                f.write("=== Caterlord 自動匯入執行報告 ===\n")
                f.write(f"執行時間: {timestamp}\n")
                f.write("="*35 + "\n\n")
                for line in report_lines:
                    f.write(line + "\n")
            
            messagebox.showinfo("執行完成", f"所有套餐已處理完畢！\n\n執行報告已儲存:\n{report_path}")
            
            # 讓程式在此無限等待，不自動關閉瀏覽器，讓您可以檢視畫面
            time.sleep(999999)
            browser.close()
            return


        # ---------------------------------------------------------
        # 【重要紀錄】: Kendo UI 下拉選單專用處理函式
        # ---------------------------------------------------------
        def select_kendo(label_text, option_text):
            try:
                # 點擊標籤後面的第一個 Kendo 輸入框 (span) 來展開選單
                page.locator(f"xpath=//label[contains(., '{label_text}')]/following::span[contains(@class, 'k-input')][1]").first.click()
                page.wait_for_timeout(500) # 等待選單動畫彈出
                # 點擊彈出的選項 (尋找當前顯示中的選單裡的對應文字)
                page.locator(f"xpath=//div[contains(@class, 'k-animation-container') and not(contains(@style, 'display: none'))]//li[contains(., '{option_text}')]").first.click()
                page.wait_for_timeout(500)
            except Exception as e:
                # 這裡不主動中斷程式，只印出警告，讓程式繼續嘗試填其他欄位
                print(f"⚠️ 選擇 {label_text} ({option_text}) 時發生錯誤: {e}")

        # ---------------------------------------------------------
        # 【重要紀錄】: 設定分店與廚房印表機
        # ---------------------------------------------------------
        # 經過實際 DOM 檢查 (inspect_result.txt)，結構如下：
        #   - 分店用 Bootstrap Panel 排列，每個 Panel 的 heading 包含 checkbox
        #   - Checkbox ID: ItemShopDetailList_{index}__isListedInShop
        #   - 廚房印表機: Kendo MultiSelect，jQuery selector = #PrinterIdList{index}
        #   - index 對照: 0=大安店, 1=Giddy, 2=台南店
        #   - Giddy 的印表機選項: Giddy_Receipt, Giddy_HOH 01, Giddy_Bar01
        #   - 台南店的印表機選項: CashierPrinter1, 煎台, 甜點, Bar外帶, Bar內用, ...
        # ---------------------------------------------------------
        SHOP_INDEX_MAP = {
            "大安店": 0,
            "Giddy": 1,
            "台南店": 2,
        }

        def setup_shop(shop_name, printer_name):
            try:
                idx = SHOP_INDEX_MAP.get(shop_name)
                if idx is None:
                    print(f"⚠️ 未知的分店名稱: {shop_name}，請在 SHOP_INDEX_MAP 中新增！")
                    return

                # 1. 打勾 checkbox
                cb_id = f"ItemShopDetailList_{idx}__isListedInShop"
                page.evaluate(f'''() => {{
                    var cb = document.getElementById("{cb_id}");
                    if (cb && !cb.checked) cb.click();
                }}''')
                
                # 2. 使用 Kendo jQuery API 直接設定 MultiSelect 的值
                #    這是最可靠的方式，完全不需要點擊 UI 元素
                page.evaluate(f'''(printerName) => {{
                    var ms = $("#PrinterIdList{idx}").data("kendoMultiSelect");
                    if (!ms) return;
                    
                    // 在資料來源中找到對應的印表機
                    var dataSource = ms.dataSource.data();
                    var targetValue = null;
                    for (var i = 0; i < dataSource.length; i++) {{
                        if (dataSource[i].PrinterName === printerName) {{
                            targetValue = dataSource[i].ShopPrinterMasterId;
                            break;
                        }}
                    }}
                    
                    if (targetValue !== null) {{
                        // 取得目前已選的值，加入新值
                        var currentValues = ms.value() || [];
                        if (currentValues.indexOf(targetValue) === -1 && currentValues.indexOf(String(targetValue)) === -1) {{
                            currentValues.push(targetValue);
                        }}
                        ms.value(currentValues);
                        ms.trigger("change");
                    }}
                }}''', printer_name)
                
                print(f"    ✅ setup_shop({shop_name}): checkbox={cb_id}, printer=PrinterIdList{idx}, value={printer_name}")
            except Exception as e:
                print(f"⚠️ 設定 {shop_name} 及其印表機時發生錯誤: {e}")

        # ---------------------------------------------------------
        # 3. 迴圈讀取所有資料並開始填寫
        # ---------------------------------------------------------
        for index, row in df.iterrows():
            item_code = str(row["商品編號"])
            item_name = str(row["商品名稱"])
            price = str(row["單價"])
            
            if not (item_code.strip().endswith('SD') or item_code.strip().endswith('SR') or item_code.strip().endswith('SF')):
                msg = f"␐跳過␑商品編號不符合規則: {item_name} ({item_code})"
                print(msg)
                report_lines.append(msg)
                continue
            
            print(f"正在處理: {item_name} ({item_code})")
            
            max_retries = 3
            for attempt in range(1, max_retries + 1):
                try:
                    # 進入商品列表頁
                    page.goto("https://hq.caterlord.com/SellableItem/")
                    page.wait_for_load_state("networkidle")

                    # 【重要紀錄】: 點擊新增按鈕 (避開 RWD 嚴格模式衝突)
                    page.locator("#ButtonAddItem").first.click()
                    page.wait_for_load_state("networkidle")

                    # --- 開始填寫表單 ---
                    
                    # 項目編碼與項目名稱
                    page.locator("#SelectedItemMaster_ItemCode").fill(item_code)
                    page.locator("#SelectedItemMaster_ItemName").fill(item_name)
                    
                    # 下拉選單 (呼叫上方寫好的 Kendo UI 專用函式)
                    select_kendo("項目種類", "一般銷售項目")
                    val_category, val_dept, val_subdept = "", "", ""
                    for col in df.columns:
                        if "分類" in col: val_category = str(row[col]).strip()
                        elif "子部門" in col: val_subdept = str(row[col]).strip()
                        elif "部門" in col: val_dept = str(row[col]).strip()
                    
                    val_category = val_category if val_category and val_category.lower() not in ['nan', 'none'] else "Beer - Taihu"
                    val_dept = val_dept if val_dept and val_dept.lower() not in ['nan', 'none'] else "Beer - Taihu"
                    val_subdept = val_subdept if val_subdept and val_subdept.lower() not in ['nan', 'none'] else "[BEV001001] BEER | Taihu"
                    
                    print(f"    -> 準備填入: 分類='{val_category}', 部門='{val_dept}', 子部門='{val_subdept}'")
                    select_kendo("分類", val_category)
                    select_kendo("部門", val_dept)
                    select_kendo("子部門", val_subdept)
                    select_kendo("按鈕樣式", "淺灰色 (細)")
                    # ---------------------------------------------------------
                    # 【重要紀錄】: 勾選「可獨立銷售及套餐項目」
                    # ---------------------------------------------------------
                    # 根據原始碼 `<label ... for="SelectedItemMaster_IsStandaloneAndSetItem">`，
                    # 得知這個 checkbox 的實際 id 為 SelectedItemMaster_IsStandaloneAndSetItem
                    # 使用 JavaScript 直接點擊，避開 viewport 錯誤，並確保只有在未勾選時才點擊
                    page.locator("#SelectedItemMaster_IsStandaloneAndSetItem").evaluate("node => { if (!node.checked) node.click(); }")
                    
                    # ---------------------------------------------------------
                    # 【重要紀錄】: 設定分店與印表機
                    # ---------------------------------------------------------
                    setup_shop("Giddy", "Giddy_Bar01")
                    code_str = str(item_code).strip()
                    name_str = str(item_name).strip()
                    if code_str.upper().startswith('S') and name_str.endswith('|S'):
                        setup_shop("台南店", "Bar外帶")
                    else:
                        setup_shop("台南店", "Bar內用")
                    # ---------------------------------------------------------
                    # 價格填寫 (Kendo NumericTextBox 特殊處理)
                    # ---------------------------------------------------------
                    page.evaluate(f'''(p) => {{
                        // 尋找所有的 Kendo 數字輸入框
                        $('input[data-role="numerictextbox"]').each(function() {{
                            var widget = $(this).data("kendoNumericTextBox");
                            if(widget) {{
                                var id = $(this).attr("id") || "";
                                var name = $(this).attr("name") || "";
                                
                                // 排除不要的特定價格欄位 (原價、套餐項目價格、成本、會員價)
                                var excludeWords = ["Original", "SetItem", "Cost", "Member", "Special"];
                                for (var i = 0; i < excludeWords.length; i++) {{
                                    if (id.includes(excludeWords[i]) || name.includes(excludeWords[i])) return; // 在 jQuery each 裡 return 相當於 continue
                                }}
                                
                                // 嘗試取得這個輸入框旁邊的文字標籤
                                var labelText = "";
                                var label = $('label[for="' + id + '"]');
                                if (label.length) labelText = label.text().trim();
                                if (!labelText) {{
                                    var parentLabel = $(this).closest('.form-group, td').find('label').first();
                                    if (parentLabel.length) labelText = parentLabel.text().trim();
                                }}
                                
                                // 如果標籤文字「剛好」等於「價格」，或者其 ID 符合各分店價格的結尾特徵
                                var isPrice = false;
                                if (labelText === "價格") isPrice = true;
                                if (id === "Price" || id === "SelectedItemMaster_Price" || id.endsWith("_Price") || id.endsWith("__Price")) isPrice = true;
                                
                                if (isPrice) {{
                                    widget.value(p);
                                    widget.trigger("change");
                                }}
                            }}
                        }});
                    }}''', float(price))
                    
                    # ---------------------------------------------------------
                    # 正式點擊儲存按鈕
                    # ---------------------------------------------------------
                    page.locator("#ButtonSave").click() 
                    
                    # 等待網路連線完成，如果網路太慢或儲存失敗會在這裡拋出 Timeout 異常
                    page.wait_for_load_state("networkidle")
                    time.sleep(1) # 停頓一下避免迴圈太快
                    
                    # 如果沒有拋出錯誤，代表這筆成功了，記錄下來
                    success_msg = f"✅ 【成功】{item_name} (編號: {item_code})"
                    if attempt > 1:
                        success_msg += f" (在第 {attempt} 次嘗試時成功)"
                    print(success_msg)
                    report_lines.append(success_msg)
                    
                    # 成功後跳出重試迴圈，前往處理下一筆資料
                    break
                    
                except Exception as e:
                    error_reason = str(e).splitlines()[0]
                    if attempt < max_retries:
                        retry_msg = f"⚠️ 【重試】{item_name} (編號: {item_code}) 第 {attempt} 次失敗: {error_reason}，準備重試..."
                        print(retry_msg)
                        time.sleep(2) # 稍微等待一下再重試
                    else:
                        error_msg = f"❌ 【徹底失敗】{item_name} (編號: {item_code}) - 已重試 {max_retries} 次仍失敗，錯誤原因: {error_reason}"
                        print(error_msg)
                        report_lines.append(error_msg)

        # ---------------------------------------------------------
        # 4. 輸出文字報告檔
        # ---------------------------------------------------------
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        report_filename = f"匯入報告_{timestamp}.txt"
        
        # 將檔案寫入與這個 Python 腳本相同的資料夾中
        script_dir = os.path.dirname(os.path.abspath(__file__))
        report_path = os.path.join(script_dir, report_filename)
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("=== Caterlord 自動匯入執行報告 ===\n")
            f.write(f"執行時間: {timestamp}\n")
            f.write("="*35 + "\n\n")
            for line in report_lines:
                f.write(line + "\n")

        # 防止瀏覽器關閉
        messagebox.showinfo("執行完畢", f"所有資料已全部處理完畢！\n\n執行報告已儲存於:\n{report_path}\n\n請至瀏覽器或報告中確認結果。")
        time.sleep(999999) 
        browser.close()

def select_file():
    file_path = filedialog.askopenfilename(
        title="選擇檔案",
        filetypes=[("資料表檔案", "*.xlsx *.xls *.csv"), ("所有檔案", "*.*")]
    )
    if file_path:
        file_label.config(text=f"📂 已選擇檔案:\n{file_path.split('/')[-1]}", fg="green")
        
        # --- 彈出選項視窗 ---
        popup = tk.Toplevel(root)
        popup.title("選擇匯入模式")
        popup.geometry("300x250")
        
        # 將 popup 視窗置中
        popup.update_idletasks()
        width = popup.winfo_width()
        height = popup.winfo_height()
        x = (popup.winfo_screenwidth() // 2) - (width // 2)
        y = (popup.winfo_screenheight() // 2) - (height // 2)
        popup.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        tk.Label(popup, text="請選擇要執行的功能：", font=("微軟正黑體", 12, "bold")).pack(pady=15)
        
        mode_var = tk.StringVar(value="全部新增")
        
        tk.Radiobutton(popup, text="1. 新增子桶", variable=mode_var, value="新增子桶", font=("微軟正黑體", 11)).pack(anchor="w", padx=70, pady=2)
        tk.Radiobutton(popup, text="2. 新增套餐", variable=mode_var, value="新增套餐", font=("微軟正黑體", 11)).pack(anchor="w", padx=70, pady=2)
        tk.Radiobutton(popup, text="3. 新增母桶", variable=mode_var, value="新增母桶", font=("微軟正黑體", 11)).pack(anchor="w", padx=70, pady=2)
        tk.Radiobutton(popup, text="4. 全部新增", variable=mode_var, value="全部新增", font=("微軟正黑體", 11)).pack(anchor="w", padx=70, pady=2)
        
        def start_now():
            selected_mode = mode_var.get()
            print(f"DEBUG: 使用者選擇了模式 -> {selected_mode}")
            popup.destroy()
            threading.Thread(target=start_automation, args=(file_path, selected_mode), daemon=True).start()
            
        tk.Button(popup, text="確認並開始執行", command=start_now, font=("微軟正黑體", 11), bg="#2196F3", fg="white", width=15).pack(pady=15)

# --- 建立 Tkinter UI 介面 ---
root = tk.Tk()
root.title("Caterlord 自動填表工具")
root.geometry("450x250")
root.eval('tk::PlaceWindow . center') 

tk.Label(root, text="請選擇要匯入的 Excel 或 CSV 檔案", font=("微軟正黑體", 14, "bold")).pack(pady=20)
file_label = tk.Label(root, text="尚未選擇檔案", fg="gray", font=("微軟正黑體", 10))
file_label.pack(pady=10)

btn = tk.Button(root, text="上傳檔案並開始自動化", command=select_file, font=("微軟正黑體", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
btn.pack(pady=10)

root.mainloop()
