# BẢN ĐỒ SỐ 22 CƠ SỞ LƯU TRÚ HỘI AN – THỰC HÀNH CHO SINH VIÊN

**Khu vực bản đồ:** Hội An, TP. Đà Nẵng (suy ra từ tên và địa chỉ của 22 cơ sở trong tệp Excel). Sheet `THONG_TIN_DU_AN` chưa khai báo khu vực/ranh giới; sinh viên cần điền phạm vi khảo sát chính thức trước khi nộp.

Bản đồ tĩnh sử dụng HTML, JavaScript, Leaflet và OpenStreetMap. Không cần cài máy chủ ứng dụng/cơ sở dữ liệu; có thể xuất bản lên GitHub Pages.

## 1. Cấu trúc bộ 3 tệp

```text
ban-do-hoi-an/
├── index.html           # Trang web bản đồ, lọc, tìm kiếm, xem ảnh và chỉ đường
├── du_lieu.js           # 22 cơ sở, hai ảnh mặc định/cơ sở, thông tin và cảnh báo
├── README.md            # Hướng dẫn này
└── anh/                 # NHÓM TỰ TẠO, hiện chưa kèm ảnh
    ├── 5S01_01.jpg      # Ảnh 1 – mặt tiền/toàn cảnh
    ├── 5S01_02.jpg      # Ảnh 2 – phòng/tiện ích
    └── 5S01_03.jpg      # Ảnh 3 – tùy chọn, chỉ thêm khi có ảnh
```

**Tệp Excel vừa gửi không kèm hình ảnh nhúng.** Ba cột K, L, M có các tên như `01_01_Mat_tien`, `01_02_Phong_nghi`, `01_03_Tien_ich`; đây là **nhãn tệp tham khảo**, chưa phải tệp ảnh, URL tải ảnh, hoặc bằng chứng có quyền sử dụng. Chúng đã được lưu trong trường `excelImageNames` để sinh viên đối chiếu.

## 2. Danh mục 22 cơ sở lưu trú

Cơ cấu theo đề bài: **7** khách sạn/resort thuộc *nhóm yêu cầu* 5 sao, **5** khách sạn nhóm 4 sao, **5** khách sạn nhóm 3 sao, **5** homestay. Tên nhóm không chứng minh từng cơ sở có giấy chứng nhận sao hợp lệ.

| Mã | Tên cơ sở theo Excel | Nhóm yêu cầu | Ảnh 1 | Ảnh 2 | Cảnh báo |
|---|---|---|---|---|---|
| `5S01` | Little Gem Hoi An Boutique Hotel & Spa | 5S | `5S01_01.jpg` | `5S01_02.jpg` | Kiểm tra |
| `5S02` | Grand Sunrise Palace Hoi An | 5S | `5S02_01.jpg` | `5S02_02.jpg` | Kiểm tra |
| `5S03` | Nolia Hoi An formerly Soluna D'annam Hoi An | 5S | `5S03_01.jpg` | `5S03_02.jpg` | Kiểm tra |
| `5S04` | Allegro Hoi An – A Little Luxury Hotel & Spa | 5S | `5S04_01.jpg` | `5S04_02.jpg` | — |
| `5S05` | Nesta Hoi An resort & Spa | 5S | `5S05_01.jpg` | `5S05_02.jpg` | Kiểm tra |
| `5S06` | Little Riverside . A Luxury Hotel And Spa | 5S | `5S06_01.jpg` | `5S06_02.jpg` | — |
| `5S07` | The Signature Hoi An | 5S | `5S07_01.jpg` | `5S07_02.jpg` | Kiểm tra |
| `4S01` | Hoi An Historic Hotel | 4S | `4S01_01.jpg` | `4S01_02.jpg` | — |
| `4S02` | Little Hoi An Boutique Hotel & Spa | 4S | `4S02_01.jpg` | `4S02_02.jpg` | Kiểm tra |
| `4S03` | Laluna Hoi An Riverside Hotel & Spa | 4S | `4S03_01.jpg` | `4S03_02.jpg` | — |
| `4S04` | Silk River Hoi An Retreat | 4S | `4S04_01.jpg` | `4S04_02.jpg` | — |
| `4S05` | Hoi An Golden Holiday Hotel & Spa | 4S | `4S05_01.jpg` | `4S05_02.jpg` | — |
| `3S01` | Hoi An Green Riverside Oasis Villa | 3S | `3S01_01.jpg` | `3S01_02.jpg` | Kiểm tra |
| `3S02` | Tomorrow Hoi An Villa | 3S | `3S02_01.jpg` | `3S02_02.jpg` | Kiểm tra |
| `3S03` | Premium CGP Hotel Hoi An | 3S | `3S03_01.jpg` | `3S03_02.jpg` | Kiểm tra |
| `3S04` | Hoi An Legend Charm Boutique Hotel | 3S | `3S04_01.jpg` | `3S04_02.jpg` | Kiểm tra |
| `3S05` | Hoi An Ivy Hotel | 3S | `3S05_01.jpg` | `3S05_02.jpg` | — |
| `HS01` | Em's House Homestay | HS | `HS01_01.jpg` | `HS01_02.jpg` | — |
| `HS02` | Hoi An Heart Lodge | HS | `HS02_01.jpg` | `HS02_02.jpg` | Kiểm tra |
| `HS03` | Jolie Villa Hoi An | HS | `HS03_01.jpg` | `HS03_02.jpg` | — |
| `HS04` | The View Homestay Hoi An | HS | `HS04_01.jpg` | `HS04_02.jpg` | Kiểm tra |
| `HS05` | Phoenix Homestay Hoi An | HS | `HS05_01.jpg` | `HS05_02.jpg` | Kiểm tra |

## 3. Cách chạy và xuất bản

1. Giải nén ZIP vào một thư mục. Đặt `index.html`, `du_lieu.js`, `README.md` cùng cấp.
2. Dùng Chrome/Edge mở `index.html` hoặc dùng VS Code → extension **Live Server** để chạy thử. **Cần Internet** để tải Leaflet và nền bản đồ OpenStreetMap.
3. Thử bấm ghim, tìm kiếm theo tên/mã hoặc tên không dấu, lọc bốn nhóm, xem chi tiết và mở Google Maps.
4. Bấm nút **▧ (Kiểm tra số ảnh)** để xem điểm nào chưa đủ ảnh thực tế và tên tệp bị thiếu.
5. Công bố: tạo repository GitHub Public, tải từng tệp đã giải nén và thư mục `anh/` (nếu đã có ảnh) lên. Trong **Settings → Pages**, chọn **Deploy from a branch**, nhánh `main`, thư mục `/(root)` và Save. Không tải ZIP lên như trang web.

Khi nhận ảnh thật từ nhóm, tạo thư mục `anh` và chép vào đúng tên tệp. Ảnh sẽ được tự động đọc khi mở popup.

## 4. Hướng dẫn ảnh minh họa: tối thiểu 2, tối đa 3

- **Ảnh 1**: mặt tiền, bảng tên, lối vào hoặc toàn cảnh, phải nhận diện đúng cơ sở.
- **Ảnh 2**: phòng nghỉ, sảnh, nhà hàng, hồ bơi hoặc tiện ích của đúng cơ sở.
- **Ảnh 3 (không bắt buộc)**: chi tiết đặc trưng như kiến trúc, cảnh quan hoặc hoạt động dịch vụ.

Dùng ảnh do nhóm tự chụp hoặc ảnh được chủ sở hữu/đơn vị quản lý cho phép công bố. Có thể mở link Google Maps hoặc website chính thức (nếu có) trong popup **để tra cứu và xin ảnh hợp lệ**. Google Maps không tự cấp quyền sao chép ảnh của người đăng; không lấy ảnh có bản quyền từ OTA/website rồi công bố như ảnh của nhóm.

Khi thêm ảnh số 3, tìm đúng mã trong `du_lieu.js` và thêm vào mảng `images`:

```js
"images": [
  {"src":"anh/5S01_01.jpg", "label":"Ảnh 1 · Mặt tiền / toàn cảnh", "credit":""},
  {"src":"anh/5S01_02.jpg", "label":"Ảnh 2 · Phòng / tiện ích", "credit":""},
  {"src":"anh/5S01_03.jpg", "label":"Ảnh 3 · Điểm đặc trưng", "credit":"Người chụp / nguồn được phép"}
]
```

Ảnh thứ 3 phải được thêm có chủ đích, không tự giả định tồn tại. Nếu ảnh là `.png` hay `.webp`, phải sửa đúng phần mở rộng trong đường dẫn `src` của cơ sở tương ứng. Không đổi tên thành `.jpg` nếu thực chất định dạng vẫn là PNG/WEBP.

**Kiểm tra số ảnh** chỉ xác định đường dẫn ảnh tải thành công, chưa xác minh nội dung ảnh đúng khách sạn hay quyền sử dụng.

## 5. Cách chỉnh sửa dữ liệu

Mở `du_lieu.js` bằng VS Code. Trong mỗi đối tượng:

- `id`: mã cơ sở cố định; `group`: 5S, 4S, 3S, HS.
- `name`, `type`, `stars`, `address`: lấy từ `DANH_SACH_22_CSLT`. Phân hạng là dữ liệu nhóm nhập, chưa xác minh độc lập.
- `lat`, `lng`: tọa độ dạng số thập phân, lấy từ Excel. **Không coi là đã kiểm tra vị trí thực địa.**
- `maps`: liên kết Google Maps nhóm cung cấp, không phải URL ảnh.
- `description`: nội dung cột N của sheet danh sách.
- `nearby`, `rooms`, `amenities`, `food`, `wellness`, `events`, `access`, `green`, `guests`, `referencePrice`, `referencePriceSource`, `highlight`, `limitations`: lấy từ sheet `KHAO_SAT_CHI_TIET`.
- `reviewNote`: chỗ hai sheet mâu thuẫn. Trước khi xóa cảnh báo, nhóm phải có bằng chứng đối chiếu nội dung và ghim.
- `images`: đường dẫn tối thiểu hai ảnh phải cung cấp; `credit`: tác giả/nguồn/giấy phép ảnh.

**Giá phòng** là ghi chép trong bảng Excel, không có bằng chứng rằng đây là giá hiện hành, vì vậy website chỉ ghi là giá tham khảo theo nguồn nhóm cung cấp.

## 6. Những mục cần đối chiếu trước khi xuất bản

Đã rà hai sheet và gắn cảnh báo cho **13 cơ sở**. Đây là các mâu thuẫn ngay trong tệp nguồn, không phải kết luận các cơ sở bị nhập sai chắc chắn:

- **`5S01` – Little Gem Hoi An Boutique Hotel & Spa:** Sheet danh sách ghi 06 Cầu Hói Muống; sheet chi tiết lại ghi tuyến Phạm Hồng Thái/Cẩm Châu. Đối chiếu lại khu vực lân cận trước khi công bố.
- **`5S02` – Grand Sunrise Palace Hoi An:** Sheet danh sách ghi 320 Cửa Đại, trong khi sheet chi tiết ghi 308 Cửa Đại. Cần đối chiếu lại địa chỉ; website chính thức hiện ghi 320 Cửa Đại.
- **`5S03` – Nolia Hoi An formerly Soluna D'annam Hoi An:** Sheet danh sách ghi 02 Cầu Hói Muống, nhưng sheet chi tiết mô tả trên đường Nguyễn Duy Hiệu. Website cơ sở ghi 01–02 Cầu Hói Muống. Nội dung chi tiết cần đối chiếu lại.
- **`5S05` – Nesta Hoi An resort & Spa:** Thuộc nhóm yêu cầu 5 sao nhưng cột hạng sao trong Excel lại ghi “Không áp dụng”; hai sheet cũng ghi 87 và 268 Cửa Đại. Website cơ sở ghi 87 Cửa Đại. Cần xác minh quyết định công nhận hạng sao.
- **`5S07` – The Signature Hoi An:** Sheet danh sách ghi 21 La Hối; sheet chi tiết lại đặt tại Hùng Vương/Thái Phiên. Website chính thức ghi 21 La Hối. Kiểm tra lại các ghi chú chi tiết.
- **`4S02` – Little Hoi An Boutique Hotel & Spa:** Sheet danh sách ghi 02 Thoại Ngọc Hầu; sheet chi tiết ghi 02 Nguyễn Thị Minh Khai. Trang chính thức ghi 02 Thoại Ngọc Hầu.
- **`3S01` – Hoi An Green Riverside Oasis Villa:** Sheet mô tả và sheet chi tiết không thống nhất về khu vực Cẩm Thanh so với Thanh Hà/Cẩm Hà; địa chỉ danh sách ghi 11 Tống Văn Sương. Cần xác định đúng cơ sở trước khi dùng nội dung chi tiết.
- **`3S02` – Tomorrow Hoi An Villa:** Sheet danh sách ghi 119 Nguyễn Tri Phương (Cẩm Nam), nhưng sheet chi tiết mô tả Cẩm Thanh, gần Rừng dừa Bảy Mẫu. Cần xác minh các ghi chú chi tiết.
- **`3S03` – Premium CGP Hotel Hoi An:** Sheet danh sách ghi 78 Trần Khát Chân, nhưng sheet chi tiết nêu Phạm Văn Đồng/Cẩm Châu. Cần đối chiếu các mô tả và tọa độ.
- **`3S04` – Hoi An Legend Charm Boutique Hotel:** Sheet danh sách ghi 121 Nguyễn Tri Phương, nhưng sheet chi tiết nêu Trần Hưng Đạo/Cẩm Phô. Cần đối chiếu bản mô tả chi tiết.
- **`HS02` – Hoi An Heart Lodge:** Sheet danh sách ghi 621/7 Hai Bà Trưng (ô địa chỉ bị lặp “621 621/7”), nhưng sheet chi tiết ghi Nguyễn Trường Tộ. Cần xác minh địa chỉ.
- **`HS04` – The View Homestay Hoi An:** Sheet danh sách ghi 28/6 Trần Hưng Đạo, nhưng sheet chi tiết mô tả Cẩm Nam/qua cầu Cẩm Nam. Cần đối chiếu phần khu vực lân cận.
- **`HS05` – Phoenix Homestay Hoi An:** Sheet danh sách ghi 06 Nguyễn Tuân, còn sheet chi tiết ghi đường Phan Đình Phùng. Cần đối chiếu thông tin địa điểm.

Các bản ghi còn lại cũng cần xác minh vị trí và hạng sao; việc không có cảnh báo không có nghĩa là đã được đối chiếu với cơ quan quản lý.

**Quan trọng:** Trong workbook nguồn, nhiều công thức kiểm tra tự động tham chiếu `#REF!` (ví dụ cột `R` của sheet danh sách và sheet tổng hợp), vì vậy không sử dụng những chỉ báo tự động đó làm bằng chứng hồ sơ đã đạt yêu cầu. Phần dữ liệu website được lấy trực tiếp từ các ô nội dung và số liệu, không dùng kết quả công thức hỏng.

### Liên kết tham khảo bổ sung

Một số website chính thức được đối chiếu bổ sung để người học kiểm tra lại địa chỉ và tìm đầu mối xin ảnh. Link chỉ là tài liệu tham khảo, **không thay cho xác nhận xếp hạng sao**:

- Little Gem Hoi An: https://www.littlegem.com.vn/
- Grand Sunrise Palace Hoi An: https://www.grandsunrisepalace.com/
- Nolia Hoi An: https://www.noliahoian.com/
- Nesta Hoi An Resort & Spa: https://nestahotel.com.vn/accommodation/nesta-hoian-resort-spa/
- The Signature Hoi An: https://www.thesignaturehoian.com/
- Little Hoi An Boutique Hotel & Spa: https://lhaghotels.com/hotels/littlehoian/contact

## 7. Checklist bài nộp sinh viên

- [ ] Phạm vi khu vực và ranh giới khảo sát đã xác định trong `THONG_TIN_DU_AN`.
- [ ] Có đúng 22 cơ sở; cơ cấu theo đề bài 7/5/5/5.
- [ ] Mỗi cơ sở có tọa độ, địa chỉ, link Maps, mô tả đã kiểm tra.
- [ ] Từng khách sạn/resort có bằng chứng phù hợp về phân hạng; homestay xác nhận đúng loại hình.
- [ ] Đã xử lý các cảnh báo về sai khác giữa hai sheet dữ liệu.
- [ ] Có ít nhất **44 ảnh thật cho 22 cơ sở** (2 ảnh/điểm), không hơn 3 ảnh mỗi điểm, đủ nguồn và quyền sử dụng.
- [ ] Bấm **▧ Kiểm tra số ảnh** đạt 22/22, xem trực tiếp để xác nhận ảnh đúng cơ sở.
- [ ] Mở thử các ghim, các nút chỉ đường/Google Maps, kiểm tra các tọa độ đúng vị trí thực.
- [ ] Chạy được trên cả máy tính và điện thoại; có URL GitHub Pages đang hoạt động.

### Công nghệ

- Leaflet: https://leafletjs.com/
- OpenStreetMap: https://www.openstreetmap.org/copyright
- Chính sách dùng lớp nền OSM: https://operations.osmfoundation.org/policies/tiles/
- GitHub Pages: https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site

Website giữ dòng ghi công OpenStreetMap; tránh thu thập hoặc công khai dữ liệu cá nhân và ảnh không được phép sử dụng.
