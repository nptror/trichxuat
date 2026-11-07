# PLAN — Tool "Lô hàng": kiểm tra, đối chiếu, theo dõi và sinh chứng từ xuất khẩu

> **Dành cho coding agent.** Đọc hết mục 1–4 trước khi viết code. Làm lần lượt Phase 0 → 3. Chỉ chuyển phase khi test và checklist "Xong khi" của phase hiện tại đã đạt.
> Không commit hay push nếu user chưa yêu cầu. Gặp quy tắc nghiệp vụ mà plan không nói rõ thì hỏi user, không tự đoán.

---

## 1. Bối cảnh nghiệp vụ

User làm chứng từ xuất khẩu nội thất khách sạn sang Mỹ. Mỗi container 40HC gom hàng của 2–3 nhà máy, mỗi nhà máy 1–2 PO. Người mua là PRIME HOSPITALITY GROUP LLC. **Mỗi cont có một thư mục chứng từ riêng.**

Mã nhà máy dùng trong tên thư mục: `SV` = SAVIMEX CORPORATION, `SKVN` = SUKAVINA COMPANY LIMITED, `NHAT HUY`/`NH` = Nhật Huy, `NHF` = chưa biết tên pháp nhân (xem mục 10).

Tên thư mục chứa cả dữ liệu, ví dụ:
`CONT 6 - 1x40HC - PO# 65395 SV + 65116, 65529 NHAT HUY + 65020, 65900 NHF - BKG EBKG19059798 đóng hàng 2`
- Số cont là `6`.
- PO → nhà máy: 65395 → SV; 65116 và 65529 → NHAT HUY; 65020 và 65900 → NHF.
- Số booking là `EBKG19059798`.
- `đóng hàng 2` nghĩa là đóng cont ngày 2.

### Luồng của một cont
1. Đặt booking với hãng tàu (COSCO, MSC…): có tàu, ETD và các hạn chót SI/VGM/CY.
2. File kế hoạch SO/CBM chia mã hàng của từng PO vào từng cont. Mốc đầy cont là 67 m³.
3. DongCont tạo DO/PKL/INV nháp.
4. Chờ nhà máy làm 3 việc: đưa hàng vào kho gom, báo KG mỗi thùng, mở tờ khai xuất.
5. Đóng cont. Ngày đóng chính là DO date; sau khi đóng mới có số cont và seal.
6. Gửi SI, VGM, ISF.
7. Forwarder phát hành B/L: 1 bản gộp và mỗi nhà máy 1 bản tách. Bản tách phải khớp tờ khai để xin C/O. Kèm theo là công văn cam kết.
8. Forwarder gửi hóa đơn phí local cho từng nhà máy.

### Quy tắc nghiệp vụ đã biết (rule trong code phải có test)
| Quy tắc | Chi tiết |
|---|---|
| Đánh số | Invoice No = `PS-{số cont 2 chữ số}.{MM}/{YYYY}`, lấy tháng/năm của DO date. Invoice date = DO date − 1 ngày. DO No = `DO -` + Invoice No. DO date = ngày đóng cont |
| VGM | VGM = tổng KG hàng + vỏ cont (tare). VGM phải ≤ MGW |
| HS | DO ghi mã 8 số ở cột P; B/L, SI, ISF ghi 6 hoặc 8 số → luôn so theo 6 số đầu (HS6) |
| Phí forwarder | Phí tính theo cont (THC, phí làm hàng, phí VGM) chia cho các nhà máy theo tỷ lệ CBM, cộng lại phải bằng mức gốc. Phí chứng từ B/L tính đủ cho mỗi B/L, không chia |
| Cảng | Long Beach = `LGB` = `USLGB`. Los Angeles (`LAX`) là **cảng khác**. Cái Mép = Vũng Tàu (cùng khu, chỉ báo ở mức thông tin). `USKCK` = Kansas City, `MEM` = Memphis |
| Tải trọng | Cảnh báo khi tổng KG hàng > 20.000 kg (ước lượng giới hạn đường bộ Mỹ). Để thành hằng số vì user sẽ xác nhận lại |

---

## 2. Hiện trạng repo (đã khảo sát)

- **Chưa có** `package.json` và `.gitignore`. Ở gốc repo có 2 thư mục lô hàng thật đang untracked, chứa giá bán và số tài khoản ngân hàng → **không bao giờ commit**.
- `index.html`: trích xuất tờ khai (file Excel ECUS) bằng SheetJS, toàn bộ code trong một `<script>`.
  - `parseWorkbook(wb)` trả về `{rows, errors}`. `rows[i].values` dùng các key của `FIELDS`: `soTK, ngayDK, loaiHinh, coQuanHQ, nguoiXK, luuKho, soKien, dvtKien, tongGross, dvtGross, maHS, moTa, sl1, dvt1, sl2, dvt2, donGia, triGia, tienTe, maXuatXu, tenXuatXu, maKemXuatXu, thongTin`.
  - Kết quả xuất ra sheet `Output`, tiêu đề cột lấy từ `FIELDS[].title` (ví dụ "Số TK", "Mã HS").
  - `declarationInfo()` ghi cứng `FROM SAVIMEX`. User đã biết và sẽ tự sửa sau, nên agent **không sửa** (`highlight-regression.cjs` cũng đang assert chuỗi này). Với tờ khai của nhà máy khác, user dán Thông tin vào Remark của DO rồi sửa tay dòng FROM (ví dụ Cont 1, DO!O23: `FROM SUKAVINA`). Vì vậy code mới **không được dựa vào dòng FROM để xác định nhà máy**: dòng này chỉ dùng cho LNT-7.
- `dongcount/DongCont.html`: từ file SO/CBM và template, tạo DO/PKL/INV bằng ExcelJS.
  - Ba khối script: `embedded-template` (`TEMPLATE_B64`), `core` (hàm thuần), `ui`.
  - `SHEET_CONFIG`: dữ liệu DO/PKL bắt đầu ở dòng 17, INV ở dòng 24.
  - `HEADER_FIELDS` gồm key và ô tương ứng trên DO: invoiceNo M6, date M7, doNo M8, doDate M9, etd D10, etaLgb D11, etaMem D12, etaLawMs D13, booking A11, contSeal A12, vessel A13, shipTo A6, contract A8. Ô trên PKL/INV xem trong code.
  - `parseCbm` trả về các dòng dạng `{id, group, po, name, code, qty, pkgs, L, W, H, auxiliary}`.
  - `fillInvoice` chỉ ghi tên, mã và công thức; không ghi QTY/UNIT/đơn giá, và **cộng cả cột đơn giá F** ở dòng TOTAL.
  - Header lưu trong `localStorage` thành một bản dùng chung cho mọi cont.
- **Test hiện có:** `highlight-regression.cjs` (cần file tờ khai fixture, không có trong repo) và `dongcount/export-regression.cjs`. Cách viết:
  - đọc file HTML, tách `<script>` bằng regex, chạy trong `node:vm`;
  - assert bằng `node:assert/strict`, đạt thì in `PASS: …`, lỗi thì đặt `process.exitCode = 1`;
  - thư viện lấy bằng `require('exceljs'|'jszip'|'xlsx')` hoặc đường dẫn qua biến môi trường.
- File DO-PKL-INV thật của Cont 1 và CONT 6 có **cùng bố cục ô** với template DongCont: header M6–M9, A11–A13, D10–D13; dòng tiêu đề DO/PKL là 15, dữ liệu từ dòng 17; INV tiêu đề dòng 22, dữ liệu từ dòng 24.

---

## 3. Quyết định kỹ thuật (đã chốt)

1. **Trang HTML một file, mở bằng `file://`** trên Edge hoặc Chrome. Không server, không bước build, **không dùng ES module**: trên `file://`, `type="module"` cùng import đường dẫn tương đối bị chặn.
2. **Thư viện lấy từ CDN, ghim phiên bản:**
   - SheetJS 0.20.3 để đọc mọi file `.xls`/`.xlsx`. ExcelJS không đọc được `.xls`, mà SI, ISF và công văn đều là `.xls`.
   - ExcelJS 4.4.0 + JSZip 3.10.1, chỉ dùng khi cần ghi xlsx có định dạng.
   - **pdf.js 3.11.174 bản UMD** (`build/pdf.min.js`); từ bản 4 trở đi chỉ còn ESM. Worker phải nạp qua blob. **Cách này đã chạy thử thành công trên `file://` với Edge 156:**
     ```js
     const src = await (await fetch('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js')).text();
     pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([src], {type: 'text/javascript'}));
     ```
3. **Đọc thư mục bằng `showDirectoryPicker()`** (đã kiểm tra: có sẵn trên `file://` ở Edge/Chrome 156). Phase 1 dùng mode `read`, Phase 2 cần `readwrite`. Dự phòng: `<input type="file" webkitdirectory>` (chỉ đọc). Bỏ qua thư mục ảnh `HÌNH LOADING`.
4. **Xử lý chữ lấy từ PDF** (đã kiểm với file thật):
   - Đọc mọi trang (tối đa 10), nối các item bằng `' '`, rồi `normalize('NFC')` và gộp khoảng trắng liên tiếp.
   - Trong hóa đơn Mentfield, pdf.js tách rời các chữ tiếng Việt mang 2 dấu (ví dụ ra "Phí x ế p d ỡ"). Vì vậy **parser bám vào nhãn tiếng Anh và con số**. Khi bắt buộc phải so chữ tiếng Việt thì dùng chuỗi "gấp": chuyển NFD, xóa `\p{M}`, xóa khoảng trắng, đổi chữ thường (ví dụ `philamhang`).
   - Định dạng số khác nhau theo mẫu chứng từ. B/L dùng dấu chấm thập phân (`11390.000 KG`, `66.870 M3`). Hóa đơn Mentfield dùng chấm phân cách nghìn, phẩy thập phân (`4.294.523` VND, `USD 164,17`). Cần hàm `parseNumber(text, 'vn'|'en')`.
5. **Không tin giá trị đã lưu sẵn trong ô công thức**, vì file do ExcelJS ghi có thể lưu kết quả là 0. Luôn tự tính:
   - CBM = L×W×H/10⁹ × số kiện;
   - KG = G.W/CTN × số kiện;
   - các dòng tổng.
6. **Đọc từng file độc lập**, mỗi file một try/catch riêng: một file lỗi không được làm hỏng kết quả của các file khác.
7. **Riêng tư và an toàn dữ liệu:**
   - Không gửi file đi đâu; mạng chỉ dùng để tải thư viện từ CDN.
   - Không sửa, xóa hay đổi tên file chứng từ.
   - Chỉ ghi `lohang.json` (từ Phase 2) và các file xuất khi user bấm nút.
8. **Nguồn sự thật là chứng từ trong thư mục.** `lohang.json` chỉ lưu những gì không đọc được từ chứng từ (xem 7.1), để không có hai nguồn dữ liệu lệch nhau.

---

## 4. Cấu trúc file đích

```
lohang/lohang.html            # UI; <script id="core"> (hàm thuần, không đụng DOM) + <script id="ui">
lohang/lohang-regression.cjs  # test Node, viết cùng kiểu dongcount/export-regression.cjs
fixtures/                     # đã gitignore — bản sao chứng từ thật để test
package.json                  # chỉ có devDependencies và script test
.gitignore
```

- Core nhận thư viện qua tham số (`XLSX`, `pdfjsLib`) để chạy được cả trong trình duyệt lẫn Node. Trong Node dùng `pdfjs-dist/legacy/build/pdf.js`; nếu có cảnh báo về font chuẩn thì bỏ qua, vì ta chỉ cần text.
- Hàm vào chính: `analyze(files, {folderName, today}) → {docs, findings}`, trong đó `files` là danh sách `{name, bytes}`.
- Viết theo phong cách code hiện có: `'use strict'`, hàm nhỏ, không framework, không thêm thư viện runtime nào ngoài pdf.js. Giao diện bằng tiếng Việt, nhãn nút bắt đầu bằng động từ, thông báo trạng thái dùng `aria-live`.

---

## 5. Phase 0 — Chuẩn bị

1. Tạo `.gitignore` gồm: `node_modules/`, `fixtures/`, `Cont */`, `CONT */`, `.DS_Store`.
2. Tạo `package.json` với `"private": true` và các devDependencies:
   - `exceljs@4.4.0`, `jszip@3.10.1`, `pdfjs-dist@3.11.174`;
   - `xlsx` cài từ `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`.

   Script `test`: `node dongcount/export-regression.cjs && node lohang/lohang-regression.cjs`.
3. Tạo fixtures bằng cách **sao chép**, không di chuyển hay sửa bản gốc:
   - `fixtures/cont1/`: mọi file trong `Cont 1 - … đóng hàng 02/Cont 1 - … đóng hàng 02/`, trừ thư mục `HÌNH LOADING`. Giữ nguyên tên file.
   - `fixtures/cont6/DO - PKL - INV - BKG EBKG19059798 - CONT 6.xlsx`: chép từ `C:\Users\MSI\Downloads\DO - PKL - INV - BKG EBKG19059798 - CONT 6 (1).xlsx`. Đây là bản nháp đang làm dở và còn ghi chú của user ở cột R–S và Sheet1.
   - `fixtures/ds-only/.DS_Store`: chép `.DS_Store` từ thư mục CONT 6 trong repo. Thư mục đó chỉ còn mỗi file này vì lỗi giải nén (đường dẫn dài hơn 260 ký tự).
4. Chạy `npm install`, rồi chạy `node dongcount/export-regression.cjs` và ghi lại kết quả **trước khi** sửa bất cứ gì.

**Xong khi:**
- `npm install` thành công;
- export-regression in PASS, hoặc lỗi sẵn có được ghi rõ và báo cho user;
- `git status` không còn hiện fixtures và thư mục lô hàng.

---

## 6. Phase 1 — Kiểm tra và đối chiếu (chỉ đọc)

Mục tiêu: user chọn thư mục của một cont; tool đọc mọi chứng từ trong đó và hiện danh sách lỗi, chỗ lệch. Phase này **không ghi gì** và **không sửa `index.html` hay `DongCont.html`**.

### 6.1 Giao diện (`lohang/lohang.html`, theo phong cách 2 tool hiện có)
- Nút chính "Mở thư mục lô hàng"; dự phòng "Chọn thư mục (chỉ đọc)".
- Bảng **Chứng từ**: tên file | loại nhận diện được | trạng thái đọc (Đọc được / Mẫu chưa hỗ trợ / Lỗi kèm lý do).
- Bảng **Kết quả**:
  - Đếm theo mức Lỗi / Cảnh báo / Thông tin / Đạt, và lọc được theo mức.
  - Mỗi phát hiện là một `<details>`: phần tóm tắt gồm tiêu đề và thông điệp. Mở ra là bảng `Chứng từ | Giá trị | Vị trí`, với vị trí là ô Excel (ví dụ `DO!G22`) hoặc số trang PDF.
  - Nhóm "Đạt" thu gọn mặc định.
- Nếu thư mục chỉ có `.DS_Store` hoặc không nhận ra file nào, hiện thông báo: "Thư mục có thể bị mất file khi giải nén (đường dẫn quá 260 ký tự). Hãy giải nén lại bằng 7-Zip vào thư mục có đường dẫn ngắn, ví dụ `D:\LH\`."

### 6.2 Nhận diện file
Nhận diện theo tên file, không phân biệt hoa thường. File xlsx có tên không khớp mẫu nào thì nhận diện theo nội dung.

| Loại | Dựa vào tên / nội dung | Định dạng |
|---|---|---|
| `doPklInv` | `DO - PKL - INV*.xlsx`, hoặc có đủ 3 sheet DO, PKL, INV | xlsx |
| `doPdf`, `pklPdf`, `invPdf` | `DO - BKG*.pdf`, `PKL - BKG*.pdf`, `INV - BKG*.pdf` | pdf |
| `booking` | tên dạng `^\d{10}\.pdf$`, hoặc trang 1 có chữ "Booking Confirmation" | pdf |
| `si` | `SI - BKG*` | xls |
| `vgm` | `VGM - BKG*.pdf` | pdf |
| `isf` | `ISF - BKG*` | xls |
| `congVan` | `Cong van cam ket*`; ưu tiên `.xls`, `.pdf` chỉ để dự phòng | xls/pdf |
| `bl` | `*Bill Of Lading*`; là bản gộp hay tách thì xác định theo nội dung | pdf |
| `fwdInvoice` | `*PRO FORMA*INVOICE*` | pdf |
| `tkOutput` | xlsx có sheet `Output` với cột `Số TK` (file xuất từ index.html) | xlsx |
| `soPlan` | `SO#*.xlsx` (Phase 1 chưa dùng) | xlsx |
| bỏ qua | ảnh, `.DS_Store`, thư mục `HÌNH LOADING` | |

### 6.3 Bộ đọc (trong core)
Mỗi bộ đọc trả về `{kind, fields, lines?, src, warnings}`, với `src[field]` là vị trí của trường đó trong chứng từ.

| Bộ đọc | Trường cần lấy | Lưu ý |
|---|---|---|
| `readDoPklInv(wb)` | Header theo `HEADER_FIELDS`. Dòng hàng DO từ dòng 17 tới `FINAL TOTAL`: po (A), name (B), descVn (C), code (D), qty (E), unit (F), pkgs (G), L/W/H (H–J), gwPerPkg (L), remark (O), hs (P). Dòng `TOTAL` đánh dấu hết nhóm của một nhà máy. INV từ dòng 24: name, code, qty (D), unit (E), price (F), và các dòng tổng | Ô A trống thì lấy PO của dòng trên; mỗi số PO mới bắt đầu một PO mới, kể cả khi vẫn trong cùng nhóm nhà máy. Dòng không có QTY nhưng có PKGS là **thùng phụ** (BOX 2/2) của dòng trên: không báo thiếu mã hay SL cho dòng này. Tự tính CBM và KG. Bỏ qua cột R–U và Sheet1 (ghi chú nháp) |
| `readSi(wb)` | shippers, consignee, notify, vessel, pod, booking, cont, seal, pkgs, kg, cbm, mgw, vgm, tare (cột `vỏ cont`), description, hs | Bảng số liệu tìm theo dòng tiêu đề `BOOKING. NO`. Tàu và POD nằm ở dòng ngay trên tiêu đề, không có nhãn |
| `readIsf(wb)` | supplier(s), seller(s), buyer, shipTo (mục 4), stuffing location (mục 5), importer, consignee, origin, hts, pol/etd, pod/eta | Tìm theo số mục `N.` ở đầu ô. Giá trị nằm sau dấu `:` trong cùng ô; nếu trống thì lấy ô bên phải hoặc dòng dưới, cho tới mục kế tiếp |
| `readCongVan(wb)` | Cảng đi, cảng đến. Bảng so "B/L chính thức" với từng "Bill n": shipper, pkgs, kg, cbm, description, hs | Mỗi dòng của cột phải có dạng `Bill n: …` |
| `readBooking(text)` | bookingNo, vessel, pol, pod, finalDest, etd, cutoff.vgm/si/cy, cargoKg, containers | Hiện chỉ hỗ trợ mẫu COSCO; mẫu khác trả về `unsupported` |
| `readBl(text)` | hbl, consignors[], consignee, vessel, pol, pod, delivery, cont, seal, pkgsText (số kiện trong mô tả, ví dụ "2000 PACKAGES"), pkgsBox (số ở ô tổng, ví dụ "200 PKGS"), kg, cbm, poList, hsList, isSplit | Là bản tách khi có chữ `A PPART OF`/`A PART OF`, hoặc chỉ có 1 consignor |
| `readFwdInvoice(text)` | invoiceNo, customerId, mbl, hbl, cont, charges[{key, usd, baseUsd, perBill}], tổng VND | Parse trên **chuỗi đã gấp**. Mỗi khoản phí bắt đầu ở mốc riêng: `thc` = `originterminalhandling`, `blFee` = `originbilloflading` (là phí perBill), `vgmFee` = `vgmsubmissionfee`, `handling` = `philamhang`. Trong đoạn của mỗi khoản: `usd(số)@` là số tiền; `baserateusd(số)` hoặc `@usd(số)/container` là mức gốc |
| `readVgm(text)` | cont, size, mgw, vgm | Dòng bảng có dạng `TGBU5033701 40HC 32,500.00 15,300.00` |
| `readDocPdf(text)` | contSeal, lấy từ dòng `CONT/SEAL NO.` | Chỉ dùng để phát hiện PDF được xuất khi chưa có số cont |
| `readTkOutput(wb)` | Các dòng tờ khai, đọc theo tiêu đề cột của FIELDS | Chỉ dùng khi thư mục có file này. Nhà máy lấy từ cột `Người XK`, không lấy từ cột `Thông tin` (dòng FROM đang ghi cứng SAVIMEX) |

**Bảng nhà máy** (đặt trong core, dễ sửa):

| Mã | Nhận diện tên pháp nhân / customer ID |
|---|---|
| SV | `SAVIMEX`, `SAVIMESGN` |
| SKVN | `SUKAVINA`, `SUKAVISGN` |
| NHAT HUY / NH | `NHAT HUY`, `NHẬT HUY` |
| NHF | chưa biết, hỏi user |

**Xác định PO thuộc nhà máy nào**, theo thứ tự ưu tiên:
1. tên thư mục (`PO# 65116, 65529 NHAT HUY + …`);
2. chữ `FOR PO# … SV + … SKVN` trong SI hoặc công văn.

Không dùng `FROM X` trong Remark: dòng này đến từ index.html (ghi cứng SAVIMEX), có thể sai nếu user quên sửa tay. Nếu dùng nó để xác định nhà máy thì LNT-7 không bao giờ bắt được lỗi đó.

### 6.4 Quy tắc kiểm tra
Mỗi rule trả về một trong ba kết quả: `pass`, `fail`, hoặc `skip` kèm lý do (ví dụ thiếu chứng từ cần thiết).

Dung sai: số kiện và số lượng phải khớp tuyệt đối; KG lệch tối đa ±0,5; CBM so sau khi làm tròn 2 chữ số thập phân; USD lệch tối đa ±0,05; VGM lệch tối đa ±1 kg.

| ID | Mức | Kiểm tra |
|---|---|---|
| TOT-1 | Lỗi | Tổng kiện, KG, CBM giống nhau giữa DO, PKL, SI, B/L gộp và công văn (cột "B/L chính thức") |
| TOT-2 | Lỗi | Tổng QTY trên INV = tổng QTY trên DO |
| SPL-1 | Lỗi | Với mỗi B/L tách: số kiện ở ô tổng, KG và CBM bằng tổng nhóm PKL của nhà máy đó, và bằng "Bill n" trong công văn |
| SPL-2 | Lỗi | Số kiện ghi trong mô tả B/L bằng số kiện ở ô tổng |
| SPL-3 | Lỗi | PO ghi trên B/L tách phải là PO của chính nhà máy đứng tên consignor |
| SPL-4 | Lỗi | Tập HS6 trên B/L tách nằm trong tập HS6 của các dòng DO thuộc nhà máy đó, và bằng HS của "Bill n" trong công văn |
| HS-1 | Cảnh báo | ISF và SI phải khai đủ mọi HS6 của các dòng hàng. Chính form ISF ghi: "All HTS Numbers … need to be reported" |
| TRN-1 | Lỗi | Tàu/chuyến giống nhau giữa booking, SI, B/L, DO và hóa đơn forwarder |
| TRN-2 | Lỗi | Số cont và seal giống nhau giữa SI, VGM, B/L và DO xlsx; bỏ qua những chỗ còn trống |
| TRN-3 | Lỗi / Cảnh báo | POD trên B/L và SI phải khớp booking (sai là Lỗi). POD trên công văn khác booking là Cảnh báo. Dùng bảng tên khác của cảng ở mục 1 |
| TRN-4 | Thông tin | POL ghi tên khác nhau nhưng cùng khu (Cái Mép / Vũng Tàu) |
| VGM-1 | Lỗi | VGM = tổng KG hàng + tare (tare lấy từ SI) |
| VGM-2 | Lỗi | VGM ≤ MGW |
| VGM-3 | Cảnh báo | Tổng KG hàng > 20.000 kg |
| FEE-1 | Lỗi | Với mỗi khoản phí tính theo cont: phần của mỗi nhà máy = mức gốc × tỷ lệ CBM của nhà máy đó, và tổng các phần = mức gốc. Mức gốc là `baseUsd` lớn nhất tìm thấy trong các hóa đơn của lô. Bỏ qua khoản `perBill` |
| NUM-1 | Lỗi | Invoice No, DO No và các ngày đúng quy tắc đánh số ở mục 1; số cont lấy từ tên thư mục |
| LNT-1 | Cảnh báo | Có dòng hàng số lượng = 0 |
| LNT-2 | Cảnh báo | Thiếu mã hàng, hoặc mã trùng nguyên văn tên hàng (mã tạm, ví dụ "QUARTZ CRATE 1") |
| LNT-3 | Cảnh báo | Thiếu UNIT, HS hoặc G.W/CTN; báo số dòng thiếu theo từng cột |
| LNT-4 | Cảnh báo | Chữ chỉ bên trái/phải trong mã (regex `([LRX])A\d{3}$`) ngược với chữ `Left`/`Right` trong tên hàng |
| LNT-5 | Cảnh báo | Hai dòng cùng tên hàng nhưng CBM mỗi thùng chênh nhau hơn 3 lần |
| LNT-6 | Cảnh báo | Trên INV: tên hàng còn ghi chú đóng gói `BOX n/m`; QTY hoặc đơn giá trống; tổng tiền = 0; ô tổng của cột đơn giá F có giá trị (cộng đơn giá với nhau là vô nghĩa) |
| LNT-7 | Cảnh báo | Remark của một PO thiếu số TK (`TK \d{12}`), hoặc ghi `FROM X` với X khác nhà máy của PO đó. Lỗi hay gặp nhất: dán Thông tin từ index.html cho PO không phải SV mà quên sửa `FROM SAVIMEX`. Remark chỉ nằm ở dòng đầu của mỗi PO, xuống dòng bằng `\r\n` |
| LNT-8 | Cảnh báo | Bản PDF của DO/PKL/INV để trống CONT/SEAL trong khi bản xlsx đã có |
| LNT-9 | Cảnh báo | ISF bỏ trống Ship-To (mục 4) |
| LNT-10 | Cảnh báo | Nghi lỗi chính tả: trong mô tả của SI/B/L/ISF/công văn có từ dài từ 6 chữ cái trở lên, không có trong tên hàng của PKL, nhưng chỉ khác 1–2 ký tự (Levenshtein) so với một từ có trong đó → gợi ý từ đúng |
| CAP-1 | Thông tin | Tổng CBM so với 67 m³; vượt 100% thì chuyển thành Cảnh báo |

### 6.5 Test với đáp án biết trước (`lohang/lohang-regression.cjs`)
Chạy `analyze` trên fixtures và so với **đúng tập phát hiện** kỳ vọng: không thiếu phát hiện nào, và không có thêm phát hiện mức Lỗi/Cảnh báo nào ngoài danh sách.

**Cont 1** (`fixtures/cont1`, folderName = `Cont 1 - 1x40HC - PO# 64834 SV + 64833 SKVN - BKG 6468146710 đóng hàng 02`). Phải phát hiện:

| Rule | Chứng từ | Giá trị đang ghi | Giá trị đúng |
|---|---|---|---|
| SPL-2 | B/L SKVN | mô tả "2000 PACKAGES" | 200 |
| SPL-3 | B/L SKVN | "PO# 64834 SV" | 64833 (SKVN) |
| SPL-4 | B/L SKVN | HS 940350 | 940320 (DO cột P; công văn Bill 2 ghi 94032090) |
| HS-1 | ISF, SI | HS6 đã khai {940350, 940320} | còn thiếu 940391, 940161, 940360 (các dòng của SV) |
| TRN-3 (Lỗi) | cả 3 B/L | POD LOS ANGELES | Long Beach (booking), LGB (SI) |
| TRN-3 (Cảnh báo) | công văn | MEMPHIS | Long Beach |
| FEE-1 | hóa đơn `_B` (SKVN) | THC 164,17; phí VGM 11,59; phí làm hàng 18,56 | khoảng 156,70; 11,19; 18,66 (SKVN chiếm 74,62% CBM). Riêng THC chênh trong khoảng 7,4–7,5 |
| FEE-1 | tổng 2 hóa đơn | THC 217,46; VGM 15,40; làm hàng 24,90 | 210; 15; 25 |
| LNT-6 | INV xlsx | ô tổng cột đơn giá = 2124 | để trống |
| LNT-8 | DO/PKL/INV pdf | CONT/SEAL trống | TGBU5033701 / 34192138 |
| LNT-9 | ISF | Ship-To trống | |
| LNT-10 | SI, B/L, ISF, công văn | HEARBOARD, NIGTSTAND | HEADBOARD, NIGHTSTAND |

Các rule sau phải **đạt**:
- TOT-1: 230 kiện / 11.390 kg / 66,87 m³; TOT-2.
- SPL-1 cho cả 2 B/L tách: SV là 30 / 2.158 / 16,97; SKVN ở ô tổng là 200 / 9.232 / 49,90.
- TRN-1; TRN-2 (TGBU5033701 / 34192138).
- VGM-1 (11.390 + 3.910 = 15.300); VGM-2 (≤ 32.500).
- NUM-1: PS-01.10/2026, ngày 01/10, DO -PS-01.10/2026, ngày 02/10.
- FEE-1 cho phần của SV: 53,29 / 3,81 / 6,34 khớp tỷ lệ 25,38%.
- LNT-7: DO!O17 (PO 64834) ghi `FROM SAVIMEX` + TK 308993421461; DO!O23 (PO 64833) ghi `FROM SUKAVINA` + TK 308991067360.

VGM-3 không được kích hoạt.

Test thêm cho LNT-7, trên bản sao trong bộ nhớ của workbook Cont 1: đổi `FROM SUKAVINA` ở DO!O23 thành `FROM SAVIMEX`. Khi đó phải có đúng 1 cảnh báo LNT-7, cho PO 64833, và nhà máy của PO 64833 vẫn phải là SKVN.

**CONT 6** (`fixtures/cont6`, folderName = tên thư mục CONT 6 ở mục 1). Fixture chỉ có bản nháp DO-PKL-INV. Phải phát hiện:

| Rule | Kỳ vọng |
|---|---|
| LNT-1 | 1 dòng có SL = 0: "Custom Wall Panel", mã MISC, DO dòng 32 |
| LNT-2 | Thiếu mã ở DO dòng 43–46 (PO 65529) và dòng 49 (PO 65900). Mã tạm ở dòng 36–42 (PO 65116) |
| LNT-3 | Cả 25/25 dòng hàng chính thiếu UNIT và HS. 30 dòng có kiện thiếu G.W/CTN: 24 dòng hàng chính có kiện + 6 thùng phụ BOX 2/2 |
| LNT-4 | Mã 19A309RA009 có chữ R nhưng tên hàng ghi "…Fridge Left" |
| LNT-5 | Cùng tên "Custom ADA Vanity - No Cubby…": 19A817XA065 thùng 0,2036 m³, còn 19A818XA066 thùng 0,0295 m³ (chênh khoảng 6,9 lần) |
| LNT-6 | INV: 6 tên hàng còn "BOX 1/2…"; QTY và đơn giá trống ở cả 25 dòng; tổng tiền = 0 |
| LNT-7 | Cả 5 PO đều chưa có số TK trong Remark |
| CAP-1 | 65,96 m³, đạt 98,5% |
| NUM-1 | **đạt**: PS-06.11/2026, 01/11/2026, DO -PS-06.11/2026, 02/11/2026 |

Các con số phải assert:
- Tổng: QTY 373, kiện 495, CBM 65,962 (±0,001).
- CBM từng PO: 65395 = 42,551; 65116 = 6,752; 65529 = 3,475; 65020 = 13,145; 65900 = 0,039.

**Không được** báo thiếu mã hay SL cho 6 dòng thùng phụ, và không được đọc các ghi chú nháp ở cột R–S hay Sheet1.

**ds-only** (`fixtures/ds-only`): chỉ ra đúng 1 thông báo "có thể mất file khi giải nén" và không có phát hiện nào khác.

### 6.6 Xong khi
- `npm test` in PASS cho cả 3 nhóm test ở trên.
- Mở `lohang/lohang.html` bằng Edge qua `file://`, chọn `fixtures/cont1`: thấy đúng các phát hiện như bảng của Cont 1; mở một phát hiện ra thấy giá trị và vị trí trên từng chứng từ.
- Không file nào trong thư mục đã chọn bị thay đổi (so thời gian sửa file trước và sau khi chạy).

---

## 7. Phase 2 — Hồ sơ lô hàng và bảng tiến độ (`lohang.json`)

### 7.1 File `lohang.json`
File nằm trong thư mục cont, ghi JSON thụt lề 2 dấu cách, ngày dạng `YYYY-MM-DD`. File chỉ lưu những gì **không đọc được từ chứng từ**:

```json
{
  "schema": 1,
  "cutoff": { "si": null, "vgm": null, "cy": null },
  "pos": [
    { "po": "65395", "warehouse": { "done": false, "eta": "2026-11-04" }, "note": "" }
  ],
  "shippers": { "SV": { "contact": "" } },
  "accepted": [ { "rule": "FEE-1", "key": "SKVN/thc", "note": "", "at": "2026-11-06" } ],
  "updatedAt": "2026-11-06T10:00:00+07:00"
}
```

Những thông tin còn lại luôn được đọc lại từ chứng từ trong thư mục, không lưu vào file này: số cont (từ tên thư mục), ngày đóng (DO date), dòng hàng, KG, tờ khai, tàu, ETD. Phải `cutoff` là ngày nhập tay, chỉ dùng khi không đọc được booking.

### 7.2 Bảng tiến độ
- **Mỗi PO một dòng**, gồm 3 cột:
  - Vào kho: ✅ đã vào, hoặc ⏳ kèm ngày dự kiến (nhập tay).
  - KG: tự đánh dấu khi mọi dòng có kiện của PO đã có G.W/CTN.
  - Tờ khai: tự đánh dấu khi Remark của PO có `TK \d{12}`, hoặc PO có trong `tkOutput`.
- **Hạn chót:** đọc từ booking (mẫu COSCO); không đọc được thì cho nhập tay. Hiện đếm ngược tới hạn SI, VGM, CY và ETD (ETD lấy từ header DO).
- **Cảnh báo** viết thành hàm thuần, nhận `today` làm tham số để test được:
  - **A1, Lỗi:** `warehouse.eta` muộn hơn ngày đóng cont.
  - **A2:** thiếu KG. Là Lỗi nếu hôm nay đã tới hạn VGM − 1 ngày, còn lại là Cảnh báo.
  - **A3, Cảnh báo:** thiếu tờ khai khi hôm nay đã tới ngày đóng − 1.
- **Nút "Soạn tin hối"** cho từng nhà máy, chép tin vào clipboard. Ví dụ:
  ```
  CONT 6 – BKG EBKG19059798 – đóng 02/11, hạn VGM 03/11
  SV – PO 65395: chưa vào kho (dự kiến 04/11), thiếu KG, thiếu tờ khai.
  ```
- **Chấp nhận phát hiện:** mỗi phát hiện của Phase 1 có nút "Chấp nhận" kèm ô ghi chú. Phát hiện đã chấp nhận được lưu vào `accepted` và ẩn khỏi danh sách, có bộ lọc "Đã chấp nhận" để xem lại.
- **Ghi và mở file:**
  - Ghi bằng `createWritable()` (trình duyệt ghi ra file tạm rồi mới thay file thật).
  - Nếu `schema` > 1 thì mở ở chế độ chỉ đọc và báo cho user.
  - Nhớ thư mục mở gần nhất trong IndexedDB; nút "Mở lại" chỉ xin quyền khi user bấm.

### 7.3 Test
- CONT 6, với `today` = 2026-10-30 và PO 65395 có eta 2026-11-04:
  - ra A1 mức Lỗi;
  - cả 5 PO thiếu KG và thiếu tờ khai;
  - tin hối gửi SV đúng nguyên văn mẫu.
- Ghi rồi đọc lại `lohang.json` thì dữ liệu giữ nguyên; các phát hiện trong `accepted` bị ẩn đúng.

**Xong khi:** test PASS. Trên Edge: mở `fixtures/cont6`, sửa trạng thái vào kho, đóng tool rồi mở lại vẫn thấy trạng thái đã sửa.

---

## 8. Phase 3 — Sinh chứng từ

1. **Sửa DongCont** (trong `core`, kèm cập nhật `export-regression.cjs`):
   - Trên INV: ghi QTY vào cột D theo dòng hàng; bỏ công thức SUM ở cột đơn giá F của dòng TOTAL; bỏ phần `BOX n/m …` khỏi tên hàng. DO và PKL giữ nguyên tên.
   - Gợi ý Invoice No, DO No và Date khi đã biết số cont (lấy từ tên file CBM, ví dụ `CONT 6 - …`) và DO date. Chỉ điền vào ô đang trống.
2. **Các khối nội dung để sao chép** trong `lohang.html`. Mỗi khối có một nút "Sao chép"; mọi giá trị tính từ chứng từ và `lohang.json`, không gõ tay:
   - **SI:** tên và địa chỉ shipper, consignee, notify, tàu, POD, booking, cont, seal, số kiện, KG, CBM, MGW, VGM, tare. Mô tả có dạng `{tên hàng} FOR PO# {po} {mã nhà máy} + …`. Danh sách HS6.
   - **VGM:** số cont, cỡ cont, MGW, VGM = KG + tare.
   - **ISF:** đủ mục 1–14, liệt kê đủ mọi HS6.
   - **Công văn cam kết:** bảng "B/L chính thức" và các "Bill n" theo từng nhà máy.
   - **Hướng dẫn B/L:** 1 bản gộp và mỗi nhà máy 1 bản tách, gom dòng hàng theo PO.

   `lohang.json` có thêm `container: { tareKg, mgwKg }` (nhập tay). Địa chỉ nhà máy lấy từ SI hoặc B/L đã có; nếu không có thì nhập tay vào `shippers`.
3. **Test với Cont 1:**
   - Khối B/L tách của SKVN phải gồm: `SUKAVINA COMPANY LIMITED`, `200 PKGS`, `9232 KGS`, `49.90 M3`, `FOR PO# 64833 SKVN`, HS `940320`.
   - Khối VGM phải ra 15.300.
4. **Làm sau**, khi user gửi template `.xlsx` của SI, VGM, ISF và công văn: điền thẳng vào template, giống cách DongCont làm với DO/PKL/INV.

---

## 9. Phase 4 — Tổng hợp nhiều cont (làm sau)

User chọn thư mục cha; tool quét các thư mục con có `lohang.json` hoặc file DO-PKL-INV và hiện một bảng gồm: cont, booking, ETD, số PO còn thiếu KG / thiếu TK / chưa vào kho, và hạn chót gần nhất. Có ô tìm kiếm theo PO, cont, booking hoặc số TK.

---

## 10. Chưa có dữ liệu hoặc cần hỏi user (không tự đoán)

- File booking MSC của CONT 6: chưa có, nên để parser MSC lại sau; trong lúc chờ thì nhập hạn chót bằng tay.
- Thư mục CONT 6 đầy đủ (sau khi giải nén lại): để thêm fixture.
- File tờ khai, hoặc file Output từ index.html, của Cont 1: để test đối chiếu B/L tách với tờ khai.
- Tên pháp nhân của NHF; xác nhận bảng mã nhà máy ↔ tên pháp nhân.
- Ngưỡng tải trọng đường bộ Mỹ (hiện mặc định 20.000 kg).
- HS trên B/L và SI nên ghi 6 số hay 8 số?

## 11. Không làm

Sửa `FROM SAVIMEX` trong index.html (user sẽ tự sửa sau), khai tờ khai (vẫn dùng ECUS), tự gửi email hoặc Zalo, OCR PDF scan, server hay đồng bộ nhiều người dùng, app desktop.
