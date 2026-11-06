from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from pathlib import Path

out=Path('PRD_Daylight_個人行程與任務管理系統.docx')
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.7);sec.bottom_margin=Inches(.7);sec.left_margin=Inches(.78);sec.right_margin=Inches(.78)
styles=doc.styles
normal=styles['Normal'];normal.font.name='Microsoft JhengHei';normal._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');normal.font.size=Pt(10);normal.font.color.rgb=RGBColor(45,49,61)
normal.paragraph_format.space_after=Pt(5);normal.paragraph_format.line_spacing=1.12
for name,size in [('Title',28),('Heading 1',17),('Heading 2',12),('Heading 3',10.5)]:
 s=styles[name];s.font.name='Microsoft JhengHei';s._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');s.font.size=Pt(size);s.font.bold=True;s.font.color.rgb=RGBColor(34,38,55);s.paragraph_format.keep_with_next=True
 if name!='Title': s.paragraph_format.space_before=Pt(12);s.paragraph_format.space_after=Pt(5)

def shade(cell, fill):
 tcPr=cell._tc.get_or_add_tcPr();shd=OxmlElement('w:shd');shd.set(qn('w:fill'),fill);tcPr.append(shd)
def borders(table,color='D8DCE6'):
 tblPr=table._tbl.tblPr; b=OxmlElement('w:tblBorders')
 for edge in ('top','left','bottom','right','insideH','insideV'):
  el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),'4');el.set(qn('w:space'),'0');el.set(qn('w:color'),color);b.append(el)
 tblPr.append(b)
def cell_margins(cell,top=90,start=100,bottom=90,end=100):
 tc=cell._tc;tcPr=tc.get_or_add_tcPr();m=OxmlElement('w:tcMar')
 for tag,val in [('top',top),('start',start),('bottom',bottom),('end',end)]:
  e=OxmlElement('w:'+tag);e.set(qn('w:w'),str(val));e.set(qn('w:type'),'dxa');m.append(e)
 tcPr.append(m)
def table(headers, rows, widths=None):
 t=doc.add_table(rows=1, cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
 if widths:
  for col,w in zip(t.columns,widths):col.width=Inches(w)
 for i,h in enumerate(headers):
  c=t.rows[0].cells[i];c.text=h;shade(c,'26334B');c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(c,120,110,120,110)
  for p in c.paragraphs:
   for r in p.runs:r.font.bold=True;r.font.color.rgb=RGBColor(255,255,255);r.font.size=Pt(9)
 trPr=t.rows[0]._tr.get_or_add_trPr();rep=OxmlElement('w:tblHeader');rep.set(qn('w:val'),'true');trPr.append(rep)
 for ridx,row in enumerate(rows):
  cells=t.add_row().cells
  for i,val in enumerate(row):
   c=cells[i];c.text=str(val);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(c)
   if ridx%2:shade(c,'F3F6FA')
   for p in c.paragraphs:
    p.paragraph_format.space_after=Pt(1)
    for r in p.runs:r.font.size=Pt(8.7)
 if widths:
  for row in t.rows:
   for i,w in enumerate(widths):row.cells[i].width=Inches(w)
 borders(t);doc.add_paragraph().paragraph_format.space_after=Pt(1)
 return t
def para(text='',boldlead=None):
 p=doc.add_paragraph()
 if boldlead and text.startswith(boldlead):p.add_run(boldlead).bold=True;p.add_run(text[len(boldlead):])
 else:p.add_run(text)
 return p
def bullets(items):
 for item in items:
  p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(item)
def numbered(items):
 for item in items:
  p=doc.add_paragraph(style='List Number');p.paragraph_format.space_after=Pt(2);p.add_run(item)
def h1(t):doc.add_heading(t,level=1)
def h2(t):doc.add_heading(t,level=2)

# Title cover, concise and content-led
p=doc.add_paragraph(style='Title');p.add_run('Daylight 個人行程與任務管理系統')
p=doc.add_paragraph();r=p.add_run('產品需求文件（PRD）');r.bold=True;r.font.size=Pt(14);r.font.color.rgb=RGBColor(102,80,205)
p=doc.add_paragraph();r=p.add_run('版本 1.0　｜　需求基線：使用者訪談　｜　日期：2026 年 10 月 30 日');r.font.size=Pt(9);r.font.color.rgb=RGBColor(105,111,128)
para('本文件定義 Daylight 多人協作行程與任務管理產品的目標使用者、核心流程、功能需求、權限與提醒原則，作為設計、開發與驗收依據。產品以個人私密資料為預設，讓使用者明確決定何時、與誰共享；首頁以「新增行程」作為最清楚、最主要的操作。')
h1('一 產品摘要')
table(['項目','定義'],[
 ('產品定位','整合個人行程、待辦任務、長期目標與多人共享的跨裝置管理工具。'),
 ('主要使用者','個人使用者，以及依邀請加入的家庭、朋友、同學或同事。協作對象可隨時新增或移除。'),
 ('核心問題','行程、待辦與長期目標分散；截止日容易遺漏；共同事項責任不清；使用者擔心私人內容被意外公開。'),
 ('產品承諾','所有內容預設私人；使用者逐項或透過共享空間明確指定分享範圍；重要提醒可持續催促至完成或手動關閉。'),
 ('主要平台','響應式網頁，支援桌面、平板與手機瀏覽器。'),
 ('本文件狀態','需求基線草案；待利害關係人確認尚未決定的服務政策與技術數值。')],[1.25,5.75])
h1('二 產品目標與成功指標')
h2('2.1 產品目標')
bullets(['讓使用者快速新增、檢視與調整每日行程。','以負責人、截止時間、優先級與狀態管理任務，讓個人及團隊清楚知道下一步。','以目標、里程碑與可執行任務連結短期行動與長期規劃。','用高可見度、可設定節奏的提醒降低遺漏，同時讓使用者保有暫停與關閉控制權。','以私人預設、明確分享、角色權限和稽核紀錄降低誤分享及誤編輯風險。'])
h2('2.2 建議成功指標（上線後建立基準再設定目標值）')
table(['指標','定義','用途'],[
 ('行程建立成功率','成功儲存行程數 ÷ 開始建立行程數','檢視建立流程是否清楚、穩定。'),
 ('任務準時完成率','截止前完成任務數 ÷ 有截止日且到期任務數','評估任務可執行性與提醒成效。'),
 ('提醒處理率','於指定期間內完成或關閉的提醒數 ÷ 發送提醒數','檢查提醒是否有用；同步觀察關閉率與退訂率，避免過度打擾。'),
 ('協作採用率','至少有一名受邀成員成功加入的共享空間數 ÷ 已建立共享空間數','衡量多人協作導入。'),
 ('私人內容誤分享事件','未經使用者明確操作而曝光私人內容的事件數','隱私與資安紅線，目標為零。')],[1.55,3.1,2.35])
h1('三 使用者與需求')
table(['使用者角色','主要需求','典型操作'],[
 ('個人使用者／建立者','集中管理個人生活與工作；預設私密；自行控制分享與提醒。','新增私人行程、任務與目標；建立共享空間；設定通知。'),
 ('空間管理者','管理空間成員、角色、邀請與共享內容。','邀請／移除成員、調整角色、維護空間設定。'),
 ('一般成員','參與共同規劃與執行，查看可存取的項目。','檢視共享行程、更新任務狀態、編輯授權內容。'),
 ('檢視者','取得資訊但不變更資料。','檢視授權的共享項目。'),
 ('指定分享對象','只存取特定行程或任務，不必加入整個空間。','接受邀請後檢視或依授權編輯單一項目。')],[1.35,2.75,2.9])
h1('四 範圍與版本策略')
h2('4.1 首版範圍')
bullets(['個人及共享空間的建立與切換；空間可邀請不同成員加入。','行程建立、編輯、刪除、日期導覽、週／月檢視、分類行事曆及搜尋。','任務清單、負責人、截止日期與時間、優先級、狀態、子任務、重複任務及逾期標示。','長期目標、期限、進度、里程碑，並可將里程碑拆成任務。','提醒中心、到期前一天與一小時提醒、每日摘要、逾期通知、重複催促與偏好設定。','依項目或共享空間控制分享範圍；管理者、成員、檢視者角色；撤銷存取。','響應式及鍵盤可操作介面、錯誤提示、空狀態、載入狀態與基本稽核紀錄。'])
h2('4.2 暫不納入首版（需另行評估）')
bullets(['與 Google／Apple／Outlook 行事曆雙向同步及複雜衝突解決。','即時聊天、視訊會議、專案文件協作及完整工作流程自動化。','由系統代發簡訊或特定通訊軟體訊息；需評估供應商、成本、同意與法規後決定。','原生 iOS／Android App；首版先以響應式網頁和瀏覽器推播評估。'])
h1('五 核心資訊架構與導航')
table(['區域','內容與用途','主要入口'],[
 ('首頁／今日總覽','今日行程、待辦摘要、即將到期、目標進度與提醒。','主導航「今日總覽」；主要行動為「新增行程」。'),
 ('行事曆','日／週／月檢視、行事曆分類、私人及共享行程。','主導航「行事曆」；新增行程按鈕固定清楚可見。'),
 ('任務管理','待辦、進行中、已完成、逾期、負責人與截止日。','主導航「任務管理」；提供新增任務。'),
 ('長期目標','目標列表、進度、期限與里程碑。','主導航「長期目標」；提供新增目標。'),
 ('提醒中心','未處理、已關閉提醒及提醒來源。','主導航及頂端通知入口。'),
 ('共享與邀請','共享空間、成員、角色、邀請及共享項目。','主導航或個人與協作入口。'),
 ('設定','提醒頻率、每日摘要時間、通知管道、安靜時段與時區。','個人選單或提醒中心。')],[1.4,3.6,2.0])
para('介面原則：畫面使用正常文件流與響應式網格，不以固定定位讓內容彼此覆蓋。桌面版善用可用寬度；較窄畫面逐步改為單欄；重要操作不得被卡片、導覽列或通知遮擋。新增行程需在首頁首屏明顯且易觸及，並與其他新增類功能清楚區隔。')
h1('六 功能需求')
requirements=[
('FR-001','首頁新增行程','首頁需將「新增行程」呈現為首要操作；按下後開啟清楚、可鍵盤操作的建立表單。','P0'),
('FR-002','行程建立與編輯','行程需支援標題、日期、開始／結束時間、時區、地點或連結、備註、分類、分享範圍與提醒。必填欄位及時間有效性需即時驗證。','P0'),
('FR-003','行事曆檢視','支援日、週、月切換及日期導覽；能區分私人與共享行事曆；選取日期後顯示對應行程。','P0'),
('FR-004','新增任務','支援標題、描述、負責人、截止日期／時間、優先級、狀態、分享範圍、提醒與子任務。','P0'),
('FR-005','任務生命週期','任務支援待辦、進行中、已完成；可編輯、排序、搜尋、篩選與標記完成；完成後停止催促。','P0'),
('FR-006','逾期處理','未完成且已超過截止時間者標記逾期；提醒中心顯示逾期通知並提供完成或手動關閉提醒操作。','P0'),
('FR-007','長期目標','支援目標名稱、說明、期限、可量化進度及多個里程碑；可將里程碑轉為任務。','P1'),
('FR-008','每日摘要','使用者可啟用每日行程摘要並設定時間與時區；摘要依當日可見行程產生。','P1'),
('FR-009','到期與重複提醒','支援到期前一天、到期前一小時、到期提醒、逾期通知及自訂催促頻率；未完成前重複提醒持續至完成或使用者手動關閉。','P0'),
('FR-010','通知管道','至少提供站內提醒；Email 與手機推播列為可設定管道。實際寄送／推播需後端、供應商、使用者同意及退訂管理。','P1'),
('FR-011','安靜時段','使用者可設安靜時段；排程提醒應延至安靜時段結束，除非使用者明確設定例外。','P1'),
('FR-012','共享空間','建立者可建立／切換空間並邀請不同 Email 成員；成員接受邀請後，依角色存取已分享內容。','P0'),
('FR-013','指定項目分享','行程及任務可指定分享對象或空間；未選擇分享時維持私人。可撤銷分享並停止後續存取。','P0'),
('FR-014','角色權限','支援管理者、成員、檢視者；權限需在伺服器端驗證，不能只靠介面隱藏操作。','P0'),
('FR-015','搜尋','可搜尋使用者有權查看的行程、任務與目標；搜尋結果不得揭露無權限內容。','P1'),
('FR-016','空間成員管理','管理者可邀請、移除成員及調整角色；重要權限變更需確認並記錄。','P1'),
('FR-017','使用者可控提醒','使用者可暫停或手動關閉特定提醒；關閉提醒不應悄悄刪除任務或行程。','P0'),
('FR-018','資料保存與多裝置','登入使用者的項目及偏好應儲存於服務端並跨裝置同步；衝突時保留明確更新時間及復原策略。','P0'),
('FR-019','基本稽核紀錄','共享資料需記錄建立、修改、完成、分享、撤銷、成員加入／移除與角色變更的操作者及時間。','P1'),
('FR-020','無障礙與響應式','支援鍵盤操作、清楚焦點、表單標籤、狀態文字與足夠色彩對比；窄螢幕內容不得互相覆蓋。','P0'),
]
table(['ID','功能','需求說明','優先級'],requirements,[.62,1.22,4.42,.72])
h1('七 主要使用流程')
h2('7.1 新增行程')
numbered(['使用者在首頁按下醒目的「新增行程」。','輸入標題、日期、起訖時間、地點／備註、分類與提醒。','選擇「私人」、「共享空間」或「指定對象」；預設為私人。若選指定對象，必須選定對象並顯示其存取範圍。','系統驗證必填欄位、結束時間晚於開始時間、可存取的共享空間及日期有效性。','儲存成功後回到原日期檢視，顯示行程及確認訊息；失敗時保留輸入並清楚說明如何修正。'])
h2('7.2 建立並協作完成任務')
numbered(['使用者新增任務並輸入負責人、截止日期／時間、優先級及提醒節奏。','使用者明確選擇分享範圍；若分享至空間，系統確認該使用者具有建立權限。','受邀成員僅能依角色查看或編輯任務；任務狀態和負責人變更需同步給有權限成員。','系統依使用者偏好在到期前提醒；逾期未完成時重複催促。','任務完成或使用者手動關閉提醒後，停止該任務的重複催促並保留任務紀錄。'])
h2('7.3 邀請成員與管理權限')
numbered(['管理者輸入受邀者 Email，選擇加入的共享空間或指定項目，再選角色。','受邀者透過有效邀請加入；邀請應具期限並可撤銷，接受前不授予資料存取權。','管理者在成員管理頁調整角色或移除成員；被移除者即時失去該空間存取權。','私人項目不因使用者加入空間而自動顯示；權限變更與撤銷寫入稽核紀錄。'])
h1('八 角色權限矩陣')
table(['操作','管理者','成員','檢視者','指定分享對象'],[
 ('查看已分享內容','是','是','是','僅受邀項目'),
 ('新增／編輯共享項目','是','是（依空間設定）','否','依單項授權'),
 ('標記任務完成','是','是','否','依單項授權'),
 ('邀請／移除成員','是','否','否','否'),
 ('調整角色／空間設定','是','否','否','否'),
 ('查看未分享私人項目','否（除非本人）','否（除非本人）','否（除非本人）','否'),
 ('撤銷自己擁有項目的分享','是','是（項目擁有者）','否','否')],[2.1,1.05,1.25,1.05,1.35])
para('權限定義需由後端執行。共用項目的建立者或擁有者可撤銷該項目分享；空間管理者可管理空間及成員，但不得因此取得成員未分享的私人資料。若未來要提供管理者代表空間成員恢復資料等例外，需獨立設計、明確告知並留存稽核紀錄。')
h1('九 資料需求與基本欄位')
table(['實體','重要欄位','關係／規則'],[
 ('使用者','user_id、名稱、Email、時區、語言、通知偏好','登入身分擁有私人項目；每項資料需有建立者與更新時間。'),
 ('行程','event_id、標題、開始／結束時間、時區、地點、備註、分類、擁有者','可屬個人或共享空間；分享權限獨立管理；時間儲存需支援時區。'),
 ('任務','task_id、標題、描述、狀態、優先級、截止時間、負責人、完成時間','可有父任務／子任務及目標里程碑關聯；完成後停止重複提醒。'),
 ('目標／里程碑','goal_id、名稱、描述、期限、進度；milestone_id、狀態','目標可包含多個里程碑；里程碑可轉換為一項或多項任務。'),
 ('共享空間／成員','space_id、名稱；member_id、角色、狀態、加入時間','成員資格屬於空間；邀請需有狀態、到期日及撤銷紀錄。'),
 ('分享授權','resource_type、resource_id、subject、permission、created_by、revoked_at','subject 可為使用者或空間；未明確授權預設拒絕。'),
 ('提醒','reminder_id、對象、觸發時間、管道、狀態、關閉原因','支援重複排程、去重、送達狀態、完成或手動關閉後停止。'),
 ('稽核紀錄','actor_id、action、resource、timestamp、變更摘要','避免記錄不必要的敏感內容；存取及保留政策待確認。')],[1.35,2.9,2.6])
h1('十 非功能需求')
table(['面向','需求'],[
 ('易用性','核心新增表單欄位清楚、錯誤訊息可理解；重要操作可由鍵盤完成；首頁資訊層級明確。'),
 ('響應式版面','支援窄至 320 CSS px 的手機寬度及一般桌面；內容採自然流動，不以覆蓋方式排版；控制項換行或堆疊時不得互相遮擋。'),
 ('效能','建議基準：一般頁面在常見網路條件下可互動時間 P75 ≤ 2.5 秒；儲存操作應提供即時載入／成功／失敗回饋。實際 SLA 由技術團隊確認。'),
 ('安全','HTTPS、身份驗證、伺服器端授權、輸入驗證、速率限制、邀請權杖防猜測、工作階段安全及依風險制定備份／復原。'),
 ('隱私','私人預設、最小權限、分享前清楚呈現對象與內容；移除權限後不再提供新讀取；明確提供資料刪除及匯出政策。'),
 ('可靠性','提醒使用佇列與冪等排程（避免重複發送）；失敗可重試並呈現送達狀態；時區及夏令時間需一致處理。'),
 ('可觀測性','記錄錯誤率、排程延遲、通知退信／拒收、同步衝突及權限拒絕；不得把敏感內容寫入一般應用程式日誌。'),
 ('無障礙','符合團隊採納的 WCAG 2.2 AA 驗收範圍；顏色不作唯一資訊線索；表單關聯標籤、焦點順序與狀態通知可用。')],[1.25,5.75])
h1('十一 提醒產品政策')
para('使用者要求醒目且像 Duolingo 一樣積極的提醒，但希望頻率可調整。提醒需讓重要事項不易忽略，同時避免未設定上限的無限轟炸。以下為建議預設與待確認政策，通知引擎不得將「手動關閉提醒」等同刪除任務。')
table(['提醒項目','建議規則','待確認事項'],[
 ('預設提醒點','行程：提前一天及一小時；任務：依截止時間提前一天及一小時。','是否對所有行程／任務預設開啟。'),
 ('重複催促','未完成時持續提醒，直到完成或使用者手動關閉；初始建議每兩小時一次。','需設定最短間隔、夜間例外及每日最高次數，以免形成濫發。'),
 ('每日摘要','使用者自選時間與時區，預設時間建議早上 8:00。','摘要是否含共享行程，以及空摘要是否寄送。'),
 ('逾期通知','截止時間過後未完成即進入逾期狀態；通知仍受使用者管道與安靜時段偏好影響。','逾期期間的頻率遞增、週末策略。'),
 ('通知管道','站內通知為基本能力；Email／推播需明確同意、偏好管理與退訂。','簡訊／通訊軟體供應商、費用及法規。'),
 ('安靜時段','使用者可設定；一般提醒延後至時段結束後。','緊急／高優先級是否允許例外。')],[1.2,3.15,2.65])
h1('十二 驗收條件')
checks=[
'首頁首屏可辨識「新增行程」為主要行動；桌面寬螢幕不把主要內容擠成狹窄左側窄欄。',
'在 320、375、768、1024、1440 CSS px 等代表寬度，首頁、行事曆、任務、目標、提醒與對話框均無文字或控制項互相重疊、被截斷或遮住。',
'未選分享範圍的新行程、任務及目標預設為私人；其他使用者無權讀取。',
'使用者可建立共享空間並邀請不同 Email；未接受邀請前受邀者無法讀取共享資料。',
'管理者、成員、檢視者的查看、編輯、完成任務、邀請與管理權限符合第八節矩陣；後端 API 同樣驗證。',
'任務可指定負責人、優先級、截止日與時間，並可完成、重新開啟及顯示逾期狀態。',
'提醒中心支援手動關閉提醒；任務完成或提醒手動關閉後，不再產生該任務的重複催促。',
'每日摘要時間與時區可設定；安靜時段內的一般提醒依政策延後。',
'共享或權限變更可在稽核記錄中追溯操作者及時間。',
'表單錯誤可理解且輸入不會無故遺失；儲存失敗時不顯示成功訊息。',
]
for c in checks:
 p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(3);p.add_run('☐  '+c)
h1('十三 假設、限制與待決策事項')
table(['待決策','目前建議','負責確認'],[
 ('登入與帳號','採 Email 驗證登入並支援帳號復原；是否提供第三方登入待決定。','產品／資安'),
 ('重複提醒上限','尊重持續催促需求，同時設可調頻率、安靜時段與合理最大頻率。','產品／使用者研究'),
 ('多人編輯衝突','先提供更新時間與明確的衝突提示；是否需即時共同編輯另評估。','產品／工程'),
 ('通知管道時程','站內提醒先行；Email、瀏覽器推播及通訊軟體分階段評估。','產品／工程／法務'),
 ('刪除與保留','定義軟刪除、復原期限、帳號刪除、資料匯出及共享資料處理。','產品／法務／資安'),
 ('時區與語言','使用者設定時區，儲存標準化時間；首版介面為繁體中文。','產品／工程'),
 ('平台責任','此 PRD 描述完整產品方向；實際多人同步、Email／推播與身份驗證需要伺服器、資料庫、排程及供應商整合。','產品／工程')],[1.45,4.15,1.4])
h1('附錄 A 名詞說明')
table(['名詞','說明'],[
 ('PRD','Product Requirements Document，產品需求文件，說明產品要解決的問題、使用者需求、功能與驗收方式。'),
 ('共享空間','使用者邀請成員加入的協作範圍；只有明確分享至空間的內容可供成員存取。'),
 ('指定項目分享','只對特定人授予單一行程或任務的存取權，不等同加入完整共享空間。'),
 ('P0／P1','優先級標籤；P0 為首版核心或必要，P1 為首版重要但可分階段交付。'),
 ('WCAG 2.2 AA','網站無障礙指引的 AA 等級，涉及鍵盤操作、對比、標籤與輔助科技支援等要求。'),
 ('冪等','同一提醒排程操作即使因重試執行多次，也只產生預期的一次有效通知。')],[1.5,5.5])

# Header/footer
for s in doc.sections:
 header=s.header.paragraphs[0];header.text='DAYLIGHT  |  產品需求文件';header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
 for r in header.runs:r.font.name='Microsoft JhengHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');r.font.size=Pt(8);r.font.color.rgb=RGBColor(115,120,136)
 footer=s.footer.paragraphs[0];footer.text='Daylight PRD  •  2026.10';footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
 for r in footer.runs:r.font.size=Pt(8);r.font.color.rgb=RGBColor(115,120,136)
# Document properties
doc.core_properties.title='Daylight 個人行程與任務管理系統 產品需求文件'
doc.core_properties.subject='多人協作行程、任務、目標與提醒需求規格'
doc.core_properties.author='Daylight Product Team'
doc.save(out)
print(out.resolve(), out.stat().st_size)
