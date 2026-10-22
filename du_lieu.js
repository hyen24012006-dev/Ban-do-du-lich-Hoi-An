/*
BẢN ĐỒ SỐ 22 CƠ SỞ LƯU TRÚ HỘI AN (ĐÀ NẴNG) – DỮ LIỆU THỰC HÀNH
Nguồn: file Excel MAU_THU_THAP_DU_LIEU_22_CSLT_THEO_KHU_VUC(AutoRecovered).xlsx
sheet DANH_SACH_22_CSLT (tên, loại, hạng, địa chỉ, tọa độ, URL Maps, mô tả)
và KHAO_SAT_CHI_TIET (phòng, tiện ích, nguồn giá, điểm nổi bật, ghi chú).

LƯU Ý VỀ TÍNH XÁC THỰC:
- Nhóm 5 sao/4 sao/3 sao phản ánh NHÓM THEO ĐỀ BÀI trong Excel; không phải
  bằng chứng cơ sở đã được cơ quan có thẩm quyền công nhận hạng sao.
- Tọa độ và nội dung trích nguyên dữ liệu nhóm nhập; chưa kiểm tra thực địa.
- Một số ô ở hai sheet mâu thuẫn; thuộc tính reviewNote mô tả việc cần đối chiếu.
- Những ô ảnh K/L/M trong Excel là TÊN ẢNH, KHÔNG PHẢI ẢNH THẬT hoặc URL.
  File được tải lên không có ảnh nhúng. 44 đường dẫn bên dưới chỉ là chỗ
  để đặt ảnh, chưa có ảnh được bổ sung. Không dùng ảnh sai cơ sở.

THÊM ẢNH:
  1) Đặt anh/5S01_01.jpg và anh/5S01_02.jpg trong thư mục anh/.
  2) Ảnh 3 KHÔNG BẮT BUỘC. Thêm {"src":"anh/5S01_03.jpg",
     "label":"Ảnh 3 · Điểm đặc trưng","credit":"Tác giả/nguồn được phép"}
     vào mảng images của cơ sở 5S01 (sau ảnh 2, thêm dấu phẩy).
  3) Dùng ảnh đã được phép công bố; ghi nguồn vào credit.
  4) Nhấn nút Kiểm tra số ảnh trên bản đồ để phát hiện ảnh thiếu.
*/
window.LODGING_PLACES = [
  {
    "id": "5S01",
    "name": "Little Gem Hoi An Boutique Hotel & Spa",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "06 Cầu Hói Muống, Hội An Đông, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8786754738004,
    "lng": 108.341721120239,
    "maps": "https://maps.app.goo.gl/pRpKBexjH4tqhTpJ7",
    "website": "https://www.littlegem.com.vn/",
    "description": "Little Gem Hoi An Boutique Hotel & Spa là khách sạn boutique tại Hội An, mang phong cách thiết kế tinh tế, kết hợp nét truyền thống và hiện đại. Khách sạn cung cấp phòng nghỉ tiện nghi, hồ bơi, nhà hàng và dịch vụ spa. Không gian được thiết kế gần gũi với thiên nhiên, tạo cảm giác thư giãn cho khách lưu trú. Với các tiện ích đa dạng, khách sạn phù hợp với khách du lịch nghỉ dưỡng, các cặp đôi và du khách muốn khám phá văn hóa địa phương.",
    "nearby": "Đường Phạm Hồng Thái / Cẩm Châu; cách Phố cổ ~800m",
    "rooms": "Deluxe Balcony, Boutique Suite, Gem Suite",
    "amenities": "Xe đạp miễn phí, Lễ tân 24/7, Đưa đón sân bay, Bàn tour",
    "food": "Nhà hàng phục vụ bữa sáng buffet / A la carte & đồ uống",
    "wellness": "Hồ bơi ngoài trời, Dịch vụ Spa & Wellness;",
    "events": "",
    "access": "Bãi đỗ xe ô tô/máy an toàn",
    "green": "Sử dụng vật liệu bài trí thân thiện môi trường, hạn chế đồ nhựa",
    "guests": "Cặp đôi, du khách quốc tế, khách nghỉ dưỡng boutique",
    "referencePrice": "1.000.000 – 2.100.000 VND",
    "referencePriceSource": "Booking.com / Agoda",
    "highlight": "Phong cách thiết kế tinh tế, ấm cúng, vị trí thuận tiện đi vào phố cổ, dịch vụ tận tình",
    "limitations": "Quy mô boutique nhỏ, diện tích hồ bơi vừa phải",
    "excelImageNames": [
      "01_01_Mat_tien",
      "01_02_Phong_nghi",
      "01_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S01_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S01_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S01_03.jpg",
    "reviewNote": "Sheet danh sách ghi 06 Cầu Hói Muống; sheet chi tiết lại ghi tuyến Phạm Hồng Thái/Cẩm Châu. Đối chiếu lại khu vực lân cận trước khi công bố."
  },
  {
    "id": "5S02",
    "name": "Grand Sunrise Palace Hoi An",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "320 Cửa Đại, Hội An Đông, Đà Nẵng 51000, Việt Nam",
    "lat": 15.8857974673526,
    "lng": 108.349954344179,
    "maps": "https://maps.app.goo.gl/pXiMVMiGYHvJNrQu9",
    "website": "https://www.grandsunrisepalace.com/",
    "description": "Grand Sunrise Palace Hoi An là khách sạn có phong cách thiết kế sang trọng, hiện đại và chú trọng trải nghiệm nghỉ dưỡng. Cơ sở cung cấp hệ thống phòng nghỉ tiện nghi cùng các dịch vụ như nhà hàng, hồ bơi và spa chăm sóc sức khỏe. Không gian được bố trí tinh tế, mang đến sự thoải mái cho khách lưu trú. Khách sạn phù hợp với khách du lịch cá nhân, gia đình và các cặp đôi có nhu cầu nghỉ dưỡng kết hợp tham quan những điểm du lịch nổi tiếng tại Hội An.",
    "nearby": "308 Cửa Đại, Cẩm Châu; cách Phố cổ ~1.8 km, Biển Cửa Đại ~2.5 km",
    "rooms": "Deluxe Balcony, Executive Suite, Grand Suite",
    "amenities": "WiFi, Dịch vụ đưa đón, Xe đạp miễn phí, Lễ tân 24/7",
    "food": "Nhà hàng Sunrise (buffet sáng, ÁÂu), Pool Bar",
    "wellness": "Hồ bơi ngoài trời, Spa & Wellness Center, Phòng Gym",
    "events": "Có phòng họp / hội nghị quy mô vừa",
    "access": "Bãi đỗ xe ô tô tại chỗ miễn phí, lối đi cho người khuyết tật",
    "green": "Hạn chế đồ nhựa dùng 1 lần, khuyến khích dùng lại khăn tắm/ga giường",
    "guests": "Gia đình, cặp đôi, khách nghỉ dưỡng",
    "referencePrice": "1.100.000 – 2.500.000 VND",
    "referencePriceSource": "Booking.com / Agoda",
    "highlight": "Kiến trúc tân cổ điển sang trọng, phòng rộng rãi, dịch vụ chuyên nghiệp",
    "limitations": "Nằm giữa khoảng Phố cổ và Biển nên cần di chuyển bằng xe đạp/Grab",
    "excelImageNames": [
      "02_01_Mat_tien",
      "02_02_Phong_nghi",
      "02_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S02_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S02_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S02_03.jpg",
    "reviewNote": "Sheet danh sách ghi 320 Cửa Đại, trong khi sheet chi tiết ghi 308 Cửa Đại. Cần đối chiếu lại địa chỉ; website chính thức hiện ghi 320 Cửa Đại."
  },
  {
    "id": "5S03",
    "name": "Nolia Hoi An formerly Soluna D'annam Hoi An",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "02 Cầu, Hói Muống, Hội An Đông, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8787513281156,
    "lng": 108.341611184656,
    "maps": "https://maps.app.goo.gl/o2ipFGMALGBeVDhaA",
    "website": "https://www.noliahoian.com/",
    "description": "Nolia Hoi An, trước đây mang tên Soluna D’Annam Hoi An, là cơ sở lưu trú phục vụ khách du lịch đến tham quan và nghỉ dưỡng. Khách sạn cung cấp phòng nghỉ tiện nghi cùng các dịch vụ hỗ trợ như lễ tân, nhà hàng và hồ bơi. Không gian lưu trú hướng đến sự thoải mái, đáp ứng nhu cầu nghỉ ngơi của nhiều nhóm khách khác nhau. Vị trí tại Hội An tạo điều kiện thuận lợi để du khách khám phá phố cổ, trải nghiệm văn hóa truyền thống và thưởng thức ẩm thực địa phương.",
    "nearby": "Đường Nguyễn Duy Hiệu, Cẩm Châu; cách Phố cổ ~1.2 km",
    "rooms": "Deluxe King/Twin, Suite",
    "amenities": "WiFi, Xe đạp miễn phí, Lễ tân 24h",
    "food": "Nhà hàng phục vụ bữa sáng và đồ uống nhẹ",
    "wellness": "Hồ bơi ngoài trời; Spa (dịch vụ cơ bản);",
    "events": "",
    "access": "Bãi đỗ xe máy/ô tô nhỏ phía trước",
    "green": "",
    "guests": "Khách du lịch tự túc, cặp đôi, nhóm bạn trẻ",
    "referencePrice": "600.000 – 1.200.000 VND",
    "referencePriceSource": "Agoda / Traveloka",
    "highlight": "Không gian ấm cúng, thiết kế hiện đại pha chút cổ điển, giá cả hợp lý",
    "limitations": "Quy mô nhỏ, tiện ích nghỉ dưỡng nâng cao còn hạn chế",
    "excelImageNames": [
      "03_01_Mat_tien",
      "03_02_Phong_nghi",
      "03_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S03_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S03_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S03_03.jpg",
    "reviewNote": "Sheet danh sách ghi 02 Cầu Hói Muống, nhưng sheet chi tiết mô tả trên đường Nguyễn Duy Hiệu. Website cơ sở ghi 01–02 Cầu Hói Muống. Nội dung chi tiết cần đối chiếu lại."
  },
  {
    "id": "5S04",
    "name": "Allegro Hoi An – A Little Luxury Hotel & Spa",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "86/2 Trần Hưng Đạo, Phường, Hội An, Đà Nẵng 51000, Việt Nam",
    "lat": 15.8809238676062,
    "lng": 108.324633357671,
    "maps": "https://maps.app.goo.gl/dzDy7WTRvmiEDytF8",
    "website": "",
    "description": "Allegro Hoi An – A Little Luxury Hotel & Spa là khách sạn mang phong cách sang trọng, kết hợp kiến trúc truyền thống Hội An với tiện nghi hiện đại. Khách sạn cung cấp phòng nghỉ chất lượng, nhà hàng, hồ bơi và dịch vụ spa chăm sóc sức khỏe. Không gian được thiết kế tinh tế, tạo cảm giác riêng tư và thư giãn cho khách lưu trú. Nhờ dịch vụ đa dạng và vị trí thuận tiện cho việc tham quan phố cổ, khách sạn phù hợp với khách quốc tế, gia đình và các cặp đôi.",
    "nearby": "86 Trần Hưng Đạo, Minh An; sát trung tâm Phố cổ (cách ~400m)",
    "rooms": "Junior Suite, Little Suite, Allegro Suite",
    "amenities": "Xe đạp miễn phí, Xe đưa đón ra bãi biển riêng An Bàng, Tour bàn",
    "food": "Nhà hàng Melody Fine Dining, Lobby Bar, Poolside Bar",
    "wellness": "Hồ bơi ngoài trời cỡ lớn, Oasis Spa cao cấp, Phòng Gym",
    "events": "Phòng họp nhỏ cho đối tác doanh nghiệp",
    "access": "Bãi đỗ xe an toàn, lối đi thuận tiện cho xe lớn",
    "green": "Sử dụng vật liệu thân thiện môi trường (tre/gỗ), giảm rác thải nhựa",
    "guests": "Khách quốc tế, khách nghỉ dưỡng cao cấp, cặp đôi trăng mật",
    "referencePrice": "2.200.000 – 4.500.000 VND",
    "referencePriceSource": "Agoda / Official Site",
    "highlight": "Vị trí đắc địa đi bộ vào phố cổ, thiết kế đậm chất di sản, dịch vụ chu đáo",
    "limitations": "Giá khá cao vào mùa cao điểm",
    "excelImageNames": [
      "04_01_Mat_tien",
      "04_02_Phong_nghi",
      "04_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S04_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S04_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S04_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "5S05",
    "name": "Nesta Hoi An resort & Spa",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Resort",
    "stars": "Không áp dụng",
    "address": "87 Cửa Đại, Hội An Đông, Đà Nẵng 564060, Việt Nam",
    "lat": 15.8932309994568,
    "lng": 108.361542244179,
    "maps": "https://maps.app.goo.gl/gCsTVgbHpRoTQ56c6",
    "website": "https://nestahotel.com.vn/accommodation/nesta-hoian-resort-spa/",
    "description": "Nesta Hoi An Resort & Spa là khu nghỉ dưỡng tại Hội An, nổi bật với không gian xanh và cảnh quan thoáng đãng. Resort cung cấp phòng nghỉ tiện nghi cùng các dịch vụ như nhà hàng, hồ bơi, spa và xe đạp phục vụ khách tham quan. Không gian nghỉ dưỡng hướng đến sự yên tĩnh, giúp du khách thư giãn sau những hoạt động khám phá. Cơ sở phù hợp với gia đình, cặp đôi và khách du lịch muốn kết hợp nghỉ dưỡng với tham quan phố cổ, khám phá cảnh quan ven sông và trải nghiệm văn hóa địa phương.",
    "nearby": "268 Cửa Đại, Cẩm Châu; sát bên sông Thu Bồn, cách Phố cổ ~2 km",
    "rooms": "Superior, Deluxe River View, Suite",
    "amenities": "Sân vườn rộng, Dịch vụ thuyền/cano, Xe đạp miễn phí",
    "food": "Nhà hàng ven sông phục vụ đặc sản Hội An & món Âu",
    "wellness": "Hồ bơi ngoài trời thoáng mát, Dịch vụ Spa chăm sóc sức khỏe;",
    "events": "Có không gian tổ chức tiệc ngoài trời / sự kiện nhỏ",
    "access": "Bãi đỗ xe rộng rãi cho cả xe khách du lịch lớn",
    "green": "Trồng nhiều cây xanh, cảnh quan tự nhiên ven sông",
    "guests": "Gia đình, nhóm đoàn du lịch, khách thích không gian yên tĩnh",
    "referencePrice": "900.000 – 1.800.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Tầm nhìn hướng sông tuyệt đẹp, không gian xanh mát, yên bình",
    "limitations": "Hơi xa khu trung tâm đi bộ, nội thất một số phòng mang phong cách truyền thống hơi cũ",
    "excelImageNames": [
      "05_01_Mat_tien",
      "05_02_Phong_nghi",
      "05_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S05_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S05_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S05_03.jpg",
    "reviewNote": "Thuộc nhóm yêu cầu 5 sao nhưng cột hạng sao trong Excel lại ghi “Không áp dụng”; hai sheet cũng ghi 87 và 268 Cửa Đại. Website cơ sở ghi 87 Cửa Đại. Cần xác minh quyết định công nhận hạng sao."
  },
  {
    "id": "5S06",
    "name": "Little Riverside . A Luxury Hotel And Spa",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "09 Đ. Phan Bội Châu, Hội An, Đà Nẵng 53100, Việt Nam",
    "lat": 15.8774297339235,
    "lng": 108.337605367544,
    "maps": "https://maps.app.goo.gl/V1HiJgkNGghxLhqL6",
    "website": "",
    "description": "Little Riverside Hoi An – A Luxury Hotel & Spa là khách sạn nằm bên sông, mang phong cách sang trọng kết hợp nét kiến trúc truyền thống và hiện đại. Khách sạn cung cấp phòng nghỉ tiện nghi, nhà hàng, hồ bơi và dịch vụ spa phục vụ nhu cầu thư giãn. Không gian ven sông tạo nên khung cảnh yên bình, thích hợp cho khách nghỉ dưỡng. Với vị trí thuận tiện để khám phá phố cổ và các điểm tham quan lân cận, khách sạn phù hợp với khách quốc tế, các cặp đôi và gia đình.",
    "nearby": "09 Phan Bội Châu, Cẩm Châu; ngay bờ sông Thu Bồn, cách Chợ Hội An ~500m",
    "rooms": "Little Suite Town/River View, Penthouse Suite",
    "amenities": "Bãi biển riêng An Bàng, Xe đưa đón bãi biển, Xe đạp miễn phí",
    "food": "Nhà hàng River Deck, Waterfront Bar",
    "wellness": "Hồ bơi vô cực tầng thượng view sông, Spa cao cấp, Gym",
    "events": "Dịch vụ sự kiện nhỏ/tiệc riêng tư",
    "access": "Có chỗ đỗ xe ô tô, không gian di chuyển thuận tiện",
    "green": "Hạn chế rác thải nhựa, ưu tiên nông sản địa phương",
    "guests": "Khách nghỉ dưỡng thư giãn, cặp đôi, khách cao cấp",
    "referencePrice": "2.500.000 – 5.000.000 VND",
    "referencePriceSource": "Agoda / Official Site",
    "highlight": "Hồ bơi vô cực view sông tuyệt đẹp, dịch vụ chuẩn 5 sao boutique, gần trung tâm",
    "limitations": "Số lượng phòng giới hạn, thường xuyên kín phòng vào mùa cao điểm",
    "excelImageNames": [
      "06_01_Mat_tien",
      "06_02_Phong_nghi",
      "06_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S06_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S06_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S06_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "5S07",
    "name": "The Signature Hoi An",
    "group": "5S",
    "category": "Khách sạn/Resort 5 sao",
    "type": "Khách sạn",
    "stars": "5 sao",
    "address": "21 La Hối, phường, Hội An, Đà Nẵng 51309, Việt Nam",
    "lat": 15.8757311549578,
    "lng": 108.324522120363,
    "maps": "https://maps.app.goo.gl/UWxpHirpxBeUmvEL7",
    "website": "https://www.thesignaturehoian.com/",
    "description": "The Signature Hoi An là khách sạn kết hợp phong cách kiến trúc hiện đại với những nét văn hóa đặc trưng của phố cổ Hội An. Cơ sở cung cấp phòng nghỉ tiện nghi, dịch vụ nhà hàng và các tiện ích phục vụ nhu cầu lưu trú của du khách. Không gian được thiết kế hài hòa, tạo cảm giác thoải mái và gần gũi với văn hóa địa phương. Nhờ vị trí thuận tiện cho việc tham quan các công trình kiến trúc cổ, thưởng thức ẩm thực và mua sắm, khách sạn phù hợp với khách du lịch văn hóa và gia đình.",
    "nearby": "Đường Hùng Vương/Thái Phiên; cách Phố cổ ~800m",
    "rooms": "Deluxe, Premium Suite, Signature Family Room",
    "amenities": "Đưa đón sân bay, Tour trải nghiệm địa phương, Xe đạp miễn phí",
    "food": "Nhà hàng phong cách ÁÂu, Quầy Bar tầng thượng/sảnh",
    "wellness": "Hồ bơi ngoài trời, Spa thư giãn, Phòng Gym",
    "events": "Phòng họp tiêu chuẩn vừa",
    "access": "Bãi đỗ xe riêng, lối vào thuận tiện cho ô tô",
    "green": "",
    "guests": "Gia đình, khách công tác kết hợp du lịch, nhóm bạn",
    "referencePrice": "1.300.000 – 2.800.000 VND",
    "referencePriceSource": "Booking.com / Traveloka",
    "highlight": "Thiết kế hiện đại pha lẫn đường nét Đông Dương, phòng mới và sạch sẽ",
    "limitations": "Tiện ích ngoài trời không quá rộng do tối ưu không gian phố",
    "excelImageNames": [
      "07_01_Mat_tien",
      "07_02_Phong_nghi",
      "07_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/5S07_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/5S07_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/5S07_03.jpg",
    "reviewNote": "Sheet danh sách ghi 21 La Hối; sheet chi tiết lại đặt tại Hùng Vương/Thái Phiên. Website chính thức ghi 21 La Hối. Kiểm tra lại các ghi chú chi tiết."
  },
  {
    "id": "4S01",
    "name": "Hoi An Historic Hotel",
    "group": "4S",
    "category": "Khách sạn 4 sao",
    "type": "Khách sạn",
    "stars": "4 sao",
    "address": "10 Trần Hưng Đạo, Hội An, Đà Nẵng 51306, Việt Nam",
    "lat": 15.8807479839573,
    "lng": 108.331244226451,
    "maps": "https://maps.app.goo.gl/dvBDqoXnHboHb9SS7",
    "website": "",
    "description": "Hoi An Historic Hotel là cơ sở lưu trú 4 sao, tọa lạc tại số 10 Trần Hưng Đạo, phường Sơn Phong, Hội An. Nơi đây cung cấp 149 phòng nghỉ với đầy đủ tiện nghi như máy điều hòa, Wi-Fi miễn phí, ban công nhìn vườn hoặc hồ bơi. Ngoài ra, cơ sở còn có hồ bơi ngoài trời, spa, nhà hàng, dịch vụ đưa đón sân bay và cho thuê xe đạp miễn phí. Với vị trí trung tâm chỉ 5 phút đi bộ đến Phố Cổ, không gian yên tĩnh và dịch vụ chu đáo, đây là lựa chọn phù hợp cho gia đình, cặp đôi và du khách muốn khám phá Hội An.",
    "nearby": "10 Trần Hưng Đạo, Sơn Phong; cách Phố cổ ~200m",
    "rooms": "Superior, Deluxe, Grand Deluxe, Suite",
    "amenities": "Khuôn viên vườn cổ thụ rộng, Sân tennis, Khu vui chơi trẻ em, Bàn tour",
    "food": "Nhà hàng Indochine, The Oaks Terrace (cà phê/bar ngoài trời)",
    "wellness": "Hồ bơi ngoài trời rất rộng, White Rose Spa, Phòng Gym",
    "events": "Trung tâm hội nghị lớn (sức chứa hàng trăm khách)",
    "access": "Bãi đỗ xe cực kỳ rộng rãi (phù hợp cả xe 45 chỗ)",
    "green": "Không gian xanh mát phủ đầy cây cổ thụ, giảm tiêu thụ năng lượng",
    "guests": "Đoàn khách MICE/Sự kiện, gia đình có trẻ nhỏ, người lớn tuổi",
    "referencePrice": "1.200.000 – 2.600.000 VND",
    "referencePriceSource": "Agoda / Official Site",
    "highlight": "Diện tích đất rộng bậc nhất ngay trung tâm phố cổ, không gian xanh thư thái",
    "limitations": "Mang phong cách kiến trúc cũ, một số hạng phòng mang nét cổ điển",
    "excelImageNames": [
      "08_01_Mat_tien",
      "08_02_Phong_nghi",
      "08_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/4S01_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/4S01_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/4S01_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "4S02",
    "name": "Little Hoi An Boutique Hotel & Spa",
    "group": "4S",
    "category": "Khách sạn 4 sao",
    "type": "Khách sạn",
    "stars": "4 sao",
    "address": "02 Thoại Ngọc Hầu, An Hội, Hội An, Đà Nẵng 550000, Việt Nam",
    "lat": 15.8762144897638,
    "lng": 108.323395157138,
    "maps": "https://maps.app.goo.gl/ey3arZe5u9w5VjqT8",
    "website": "https://lhaghotels.com/hotels/littlehoian/contact",
    "description": "Little Hoi An Boutique Hotel & Spa là cơ sở lưu trú 4 sao, tọa lạc tại số 02 Thoại Ngọc Hầu, Hội An. Nơi đây có 30 phòng nghỉ trang nhã với máy điều hòa, Wi-Fi miễn phí, phòng tắm riêng và ban công nhìn sông Hoài. Cơ sở còn có hồ bơi trong nhà, spa, nhà hàng phục vụ ẩm thực Việt và Tây, lớp học nấu ăn cùng trung tâm thể dục. Chỉ cách Phố Cổ khoảng 300m, không gian ấm áp, nhân viên thân thiện và bữa sáng đa dạng, rất phù hợp cho du khách cá nhân, cặp đôi và những ai yêu thích sự gần gũi.",
    "nearby": "02 Nguyễn Thị Minh Khai; cách Chùa Cầu ~300m",
    "rooms": "Little Deluxe, Little Suite",
    "amenities": "Xe đạp miễn phí, Xe bus đưa đón bãi biển An Bàng, Bàn tour",
    "food": "Nhà hàng Gochudang / Gotique (Á  Âu & đặc sản)",
    "wellness": "Hồ bơi trong nhà/bán ngoài trời, Spa, Gym nhỏ",
    "events": "",
    "access": "Bãi đỗ xe nhỏ (do gần khu vực đi bộ)",
    "green": "Áp dụng mô hình thân thiện môi trường của Little Hoi An Group",
    "guests": "Khách du lịch nước ngoài, cặp đôi thích đi bộ khám phá phố cổ",
    "referencePrice": "1.100.000 – 2.200.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Ngay sát Chùa Cầu, nhân viên thân thiện, thiết kế hoài cổ",
    "limitations": "Hồ bơi diện tích khiêm tốn",
    "excelImageNames": [
      "09_01_Mat_tien",
      "09_02_Phong_nghi",
      "09_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/4S02_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/4S02_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/4S02_03.jpg",
    "reviewNote": "Sheet danh sách ghi 02 Thoại Ngọc Hầu; sheet chi tiết ghi 02 Nguyễn Thị Minh Khai. Trang chính thức ghi 02 Thoại Ngọc Hầu."
  },
  {
    "id": "4S03",
    "name": "Laluna Hoi An Riverside Hotel & Spa",
    "group": "4S",
    "category": "Khách sạn 4 sao",
    "type": "Khách sạn",
    "stars": "4 sao",
    "address": "12 Nguyễn Du, Hội An, Đà Nẵng 550000, Việt Nam",
    "lat": 15.8767473456572,
    "lng": 108.322204282271,
    "maps": "https://maps.app.goo.gl/CZTi2hxoZ5WNqqmB8",
    "website": "",
    "description": "Laluna Hoi An Riverside Hotel & Spa là cơ sở lưu trú 4 sao, tọa lạc tại số 12 Nguyễn Du, phường Cẩm Phô, Hội An. Nơi đây có 48 phòng nghỉ hiện đại với máy điều hòa, Wi-Fi miễn phí, ban công nhìn sông và phòng gia đình rộng rãi. Cơ sở có hồ bơi ngoài trời, spa chuyên nghiệp, nhà hàng với ẩm thực Á-Âu đa dạng và cho thuê xe đạp miễn phí. Chỉ 3 phút đi bộ đến Chùa Cầu, vị trí tuyệt vời, phòng sạch sẽ và dịch vụ chu đáo, phù hợp với mọi đối tượng du khách đến Hội An",
    "nearby": "12 Nguyễn Du, Cẩm Phô; ven sông Hoài, cách Chùa Cầu ~400m",
    "rooms": "Superior Window, Deluxe River View, Suite Balcony",
    "amenities": "Quầy bar tầng trệt, Dịch vụ hỗ trợ du lịch, Mượn xe đạp",
    "food": "Nhà hàng Laluna (view sông Hoài)",
    "wellness": "Hồ bơi ngoài trời mát mẻ, Laluna Spa;",
    "events": "",
    "access": "Có chỗ đỗ xe ô tô gần khách sạn",
    "green": "",
    "guests": "Khách du lịch tự túc, cặp đôi, gia đình nhỏ",
    "referencePrice": "900.000 – 1.900.000 VND",
    "referencePriceSource": "Agoda / Traveloka",
    "highlight": "Vị trí đắc địa nhìn ra sông Hoài, buổi tối ngắm đèn lồng rất đẹp, đi bộ vào trung tâm nhanh",
    "limitations": "Khu vực đường Nguyễn Du có thể hơi đông đúc vào giờ cao điểm buổi tối",
    "excelImageNames": [
      "10_01_Mat_tien",
      "10_02_Phong_nghi",
      "10_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/4S03_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/4S03_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/4S03_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "4S04",
    "name": "Silk River Hoi An Retreat",
    "group": "4S",
    "category": "Khách sạn 4 sao",
    "type": "Khách sạn",
    "stars": "4 sao",
    "address": "12 Huyền Trân Công Chúa, Hội An Đông, Đà Nẵng 51000, Việt Nam",
    "lat": 15.8763274759491,
    "lng": 108.3470864899,
    "maps": "https://maps.app.goo.gl/c6gTKiHwVg41F2oZ9",
    "website": "",
    "description": "Silk River Hoi An Retreat là cơ sở lưu trú 4 sao, tọa lạc tại số 12 Huyền Trân Công Chúa, Hội An. Nơi đây cung cấp 50 phòng nghỉ thiết kế hài hòa giữa truyền thống và hiện đại, có máy điều hòa, Wi-Fi miễn phí, ban công rộng nhìn sông hoặc hồ bơi. Cơ sở còn có hồ bơi ngoài trời giữa sân vườn xanh mát, spa, nhà hàng view sông và dịch vụ cho thuê xe đạp. Không gian thanh bình, yên tĩnh, dễ di chuyển đến Phố Cổ và biển An Bàng, rất phù hợp cho kỳ nghỉ dưỡng dài ngày cùng gia đình hoặc bạn bè.",
    "nearby": "Khu vực Cẩm Châu/Cẩm Nam (ven sông Thu Bồn); cách Phố cổ ~1.5 km",
    "rooms": "Deluxe River View, Suite Balcony, Family Room",
    "amenities": "Xe đạp miễn phí, Sân vườn thư giãn, Shuttle bus theo lịch",
    "food": "Nhà hàng ven sông phục vụ buffet sáng & A la carte",
    "wellness": "Hồ bơi ngoài trời nhìn ra sông, Dịch vụ Spa/Massage;",
    "events": "",
    "access": "Chỗ đỗ xe ô tô an toàn",
    "green": "Tối đa hóa ánh sáng tự nhiên, không gian xanh thoáng đãng",
    "guests": "Khách nghỉ dưỡng thích sự riêng tư, thư thái",
    "referencePrice": "800.000 – 1.600.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Không gian yên tĩnh, view sông thơ mộng, không khí trong lành",
    "limitations": "Cần di chuyển bằng xe đạp/xe máy nếu muốn vào trung tâm ăn đêm",
    "excelImageNames": [
      "11_01_Mat_tien",
      "11_02_Phong_nghi",
      "11_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/4S04_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/4S04_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/4S04_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "4S05",
    "name": "Hoi An Golden Holiday Hotel & Spa",
    "group": "4S",
    "category": "Khách sạn 4 sao",
    "type": "Khách sạn",
    "stars": "4 sao",
    "address": "467 Đ. Hai Bà Trưng, Hội An Tây, Đà Nẵng, Việt Nam",
    "lat": 15.8876405040427,
    "lng": 108.32831408042,
    "maps": "https://maps.app.goo.gl/AhM9Vi3Fe2QiFFbP8",
    "website": "",
    "description": "Hoi An Golden Holiday Hotel & Spa là cơ sở lưu trú 4 sao, nằm trên đường Hai Bà Trưng, Hội An. Nơi đây có 60 phòng nghỉ sáng sủa với máy điều hòa, Wi-Fi miễn phí, tiện nghi đầy đủ và view hồ bơi hoặc vườn. Cơ sở có hồ bơi ngoài trời kèm quầy bar, spa, phòng gym hoạt động 24 giờ, nhà hàng và cho thuê xe đạp miễn phí. Nằm giữa Phố Cổ và bãi biển, giá cả hợp lý cùng nhiều ưu đãi hấp dẫn, là lựa chọn lý tưởng cho du khách và gia đình khi đến Hội An.",
    "nearby": "467 Hai Bà Trưng, Tân An; cách Phố cổ ~1.2 km",
    "rooms": "Deluxe Balcony, Suite, Family Connecting",
    "amenities": "Cho thuê xe máy, Xe đạp miễn phí, Đưa đón sân bay",
    "food": "Nhà hàng Golden Holiday (món ăn đa dạng)",
    "wellness": "Hồ bơi ngoài trời, Dịch vụ Spa chăm sóc sức khỏe;",
    "events": "",
    "access": "Bãi đỗ xe ô tô phía trước khách sạn",
    "green": "",
    "guests": "Gia đình, nhóm bạn, khách du lịch tiết kiệm",
    "referencePrice": "650.000 – 1.300.000 VND",
    "referencePriceSource": "Agoda / Traveloka",
    "highlight": "Giá cả rất hợp lý, phòng sạch sẽ, nằm trên trục đường chính thuận tiện",
    "limitations": "Khách sạn tuyến trung cấp, không quá nhiều tiện ích xa hoa",
    "excelImageNames": [
      "12_01_Mat_tien",
      "12_02_Phong_nghi",
      "12_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/4S05_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/4S05_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/4S05_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "3S01",
    "name": "Hoi An Green Riverside Oasis Villa",
    "group": "3S",
    "category": "Khách sạn 3 sao",
    "type": "Khách sạn",
    "stars": "3 sao",
    "address": "11 Tống Văn Sương, Hội An Đông, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8827078273477,
    "lng": 108.352234620897,
    "maps": "https://maps.app.goo.gl/KJU7mQYoUUEQeKsE6",
    "website": "",
    "description": "Hoi An Green Riverside Oasis Villa nằm tại khu vực Cẩm Thanh, Hội An, có không gian gần gũi với thiên nhiên. Cơ sở cung cấp phòng nghỉ được trang bị máy điều hòa, phòng tắm riêng và Wi-Fi miễn phí. Các tiện ích được giới thiệu gồm hồ bơi ngoài trời, khu vườn, nhà hàng và xe đạp miễn phí dành cho khách. Không gian tương đối yên tĩnh, phù hợp với du khách muốn nghỉ ngơi và khám phá khu vực địa phương. Do không nằm ngay trung tâm phố cổ, khách có thể cần phương tiện di chuyển khi tham quan.",
    "nearby": "Khu vực Thanh Hà / Cẩm Hà (gần sông); cách Phố cổ ~2.5 km",
    "rooms": "Villa Room View Garden/Pool, Family Suite",
    "amenities": "Sân vườn xanh mát, Bếp chung/dụng cụ BBQ, Xe đạp miễn phí",
    "food": "Phục vụ bữa sáng theo yêu cầu, nước uống tại hồ bơi",
    "wellness": "Hồ bơi ngoài trời bao quanh bởi vườn cây;  ;",
    "events": "Tiệc ngoài trời / BBQ nhóm nhỏ",
    "access": "Chỗ đỗ xe ô tô thoải mái",
    "green": "Môi trường nhiều cây xanh, kiến trúc mở hòa hợp thiên nhiên",
    "guests": "Gia đình, nhóm bạn muốn ở không gian villa biệt lập",
    "referencePrice": "500.000 – 1.100.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Không gian biệt thự xanh mát, yên tĩnh, tạo cảm giác như ở nhà",
    "limitations": "Xa trung tâm phố cổ hơn so với các khách sạn nội thị",
    "excelImageNames": [
      "13_01_Mat_tien",
      "13_02_Phong_nghi",
      "13_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/3S01_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/3S01_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/3S01_03.jpg",
    "reviewNote": "Sheet mô tả và sheet chi tiết không thống nhất về khu vực Cẩm Thanh so với Thanh Hà/Cẩm Hà; địa chỉ danh sách ghi 11 Tống Văn Sương. Cần xác định đúng cơ sở trước khi dùng nội dung chi tiết."
  },
  {
    "id": "3S02",
    "name": "Tomorrow Hoi An Villa",
    "group": "3S",
    "category": "Khách sạn 3 sao",
    "type": "Khách sạn",
    "stars": "3 sao",
    "address": "119 Nguyễn Tri Phương, Hội An, Đà Nẵng, Việt Nam",
    "lat": 15.8724969977699,
    "lng": 108.336565268245,
    "maps": "https://maps.app.goo.gl/GQttpxLNXpP6gyij6",
    "website": "",
    "description": "Tomorrow Hoi An Villa tọa lạc tại khu vực Cẩm Nam, cách phố cổ Hội An khoảng 500 m theo thông tin giới thiệu trực tuyến. Phòng nghỉ có máy điều hòa, phòng tắm riêng, tivi và Wi-Fi miễn phí. Các tiện ích nổi bật gồm hồ bơi ngoài trời, khu vườn, sân hiên, nhà hàng và dịch vụ cho thuê xe đạp. Vị trí tương đối thuận tiện để tham quan phố cổ, đồng thời mang lại không gian nghỉ ngơi yên tĩnh hơn so với khu vực trung tâm đông đúc.",
    "nearby": "Khối 2, Cẩm Thanh (gần Rừng dừa Bảy Mẫu); cách Phố cổ ~3 km",
    "rooms": "Deluxe King/Twin View Cánh Đồng/Hồ Bơi",
    "amenities": "Xe đạp miễn phí, Hỗ trợ đặt tour Rừng dừa, Đưa đón sân bay",
    "food": "Bữa sáng gọi món (Set menu/A la carte), Phục vụ giải khát",
    "wellness": "Hồ bơi ngoài trời giữa lòng villa;  ;",
    "events": "",
    "access": "Có chỗ đỗ ô tô riêng",
    "green": "Tận dụng khung cảnh thiên nhiên đồng quê lúa xanh/rừng dừa",
    "guests": "Khách nước ngoài thích trải nghiệm đồng quê, khách nghỉ dưỡng tĩnh lặng",
    "referencePrice": "450.000 – 900.000 VND",
    "referencePriceSource": "Agoda / Booking.com",
    "highlight": "Tầm nhìn ra đồng lúa/thiên nhiên xanh mát, chủ nhà rất thân thiện, giá rẻ",
    "limitations": "Xa trung tâm mua sắm, thích hợp nghỉ ngơi hơn là di chuyển liên tục",
    "excelImageNames": [
      "14_01_Mat_tien",
      "14_02_Phong_nghi",
      "14_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/3S02_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/3S02_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/3S02_03.jpg",
    "reviewNote": "Sheet danh sách ghi 119 Nguyễn Tri Phương (Cẩm Nam), nhưng sheet chi tiết mô tả Cẩm Thanh, gần Rừng dừa Bảy Mẫu. Cần xác minh các ghi chú chi tiết."
  },
  {
    "id": "3S03",
    "name": "Premium CGP Hotel Hoi An",
    "group": "3S",
    "category": "Khách sạn 3 sao",
    "type": "Khách sạn",
    "stars": "3 sao",
    "address": "78 Trần Khát Chân, Hội An Đông, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8817805623184,
    "lng": 108.339967413492,
    "maps": "https://maps.app.goo.gl/CneS1qX56K5eDTbr8",
    "website": "",
    "description": "Premium CGP Hotel Hoi An nằm trên đường Trần Khát Chân, thuộc khu vực Cẩm Châu, Hội An. Khách sạn cung cấp phòng nghỉ có máy điều hòa, phòng tắm riêng, tivi và Wi-Fi miễn phí. Những tiện ích được giới thiệu gồm hồ bơi, khu vườn, sân hiên, nhà hàng và quầy bar. Ngoài ra, cơ sở có phòng gia đình, dịch vụ phòng và chỗ đỗ xe. Các tiện ích này đáp ứng nhu cầu nghỉ ngơi, ăn uống và thư giãn của khách lưu trú. Khách có thể sử dụng phương tiện cá nhân hoặc dịch vụ di chuyển để khám phá các điểm tham quan.",
    "nearby": "Đường Phạm Văn Đồng / Cẩm Châu; cách Phố cổ ~1.8 km",
    "rooms": "Deluxe City View, Premium Suite",
    "amenities": "Wifi tốc độ cao, Lễ tân 24/7, Xe đạp miễn phí",
    "food": "Nhà hàng phục vụ bữa sáng buffet / A la carte",
    "wellness": "Hồ bơi ngoài trời; Spa cơ bản;",
    "events": "Phòng họp nhỏ",
    "access": "Bãi đỗ xe ô tô an toàn",
    "green": "",
    "guests": "Khách công tác, cặp đôi, nhóm bạn",
    "referencePrice": "700.000 – 1.400.000 VND",
    "referencePriceSource": "Traveloka / Agoda",
    "highlight": "Khách sạn mới, trang thiết bị hiện đại, sạch sẽ",
    "limitations": "Ít cây xanh khuôn viên so với các resort ven sông",
    "excelImageNames": [
      "15_01_Mat_tien",
      "15_02_Phong_nghi",
      "15_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/3S03_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/3S03_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/3S03_03.jpg",
    "reviewNote": "Sheet danh sách ghi 78 Trần Khát Chân, nhưng sheet chi tiết nêu Phạm Văn Đồng/Cẩm Châu. Cần đối chiếu các mô tả và tọa độ."
  },
  {
    "id": "3S04",
    "name": "Hoi An Legend Charm Boutique Hotel",
    "group": "3S",
    "category": "Khách sạn 3 sao",
    "type": "Khách sạn",
    "stars": "3 sao",
    "address": "121 Nguyễn Tri Phương, Hội An, Đà Nẵng 560000, Việt Nam",
    "lat": 15.872591394403,
    "lng": 108.336233128302,
    "maps": "https://maps.app.goo.gl/XosBq82tc8wHM8LL6",
    "website": "",
    "description": "Hoi An Legend Charm Boutique Hotel nằm tại khu vực Cẩm Nam, gần khu phố cổ và chợ đêm Hội An. Theo thông tin giới thiệu trực tuyến, khách sạn có các phòng nghỉ tương đối rộng; một số phòng có ban công hoặc tầm nhìn ra sông và khu vườn. Các tiện ích được giới thiệu gồm hồ bơi, nhà hàng, Wi-Fi miễn phí, dịch vụ spa và cho thuê xe đạp. Khu vực sân vườn tạo thêm không gian thư giãn cho khách. Cơ sở phù hợp với du khách muốn kết hợp tham quan phố cổ với nghỉ ngơi trong thời gian lưu trú",
    "nearby": "Đường Trần Hưng Đạo / Cẩm Phô; cách Phố cổ ~700m",
    "rooms": "Superior, Deluxe Balcony, Legend Suite",
    "amenities": "Cho thuê xe máy, Xe đạp miễn phí, Tư vấn tour",
    "food": "Nhà hàng nhỏ phục vụ ăn sáng",
    "wellness": "Hồ bơi ngoài trời khiêm tốn;  ;",
    "events": "",
    "access": "Bãi đỗ xe máy/ô tô nhỏ",
    "green": "",
    "guests": "Khách du lịch tự túc, backpackers, cặp đôi",
    "referencePrice": "550.000 – 1.100.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Vị trí tiện lợi đi bộ, phong cách thiết kế đậm chất boutique hoài cổ",
    "limitations": "Hồ bơi nhỏ, không gian chung hạn chế",
    "excelImageNames": [
      "16_01_Mat_tien",
      "16_02_Phong_nghi",
      "16_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/3S04_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/3S04_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/3S04_03.jpg",
    "reviewNote": "Sheet danh sách ghi 121 Nguyễn Tri Phương, nhưng sheet chi tiết nêu Trần Hưng Đạo/Cẩm Phô. Cần đối chiếu bản mô tả chi tiết."
  },
  {
    "id": "3S05",
    "name": "Hoi An Ivy Hotel",
    "group": "3S",
    "category": "Khách sạn 3 sao",
    "type": "Khách sạn",
    "stars": "3 sao",
    "address": "01 Đào Duy Từ, Hội An, Đà Nẵng 51000, Việt Nam",
    "lat": 15.8771069540416,
    "lng": 108.323001084122,
    "maps": "https://maps.app.goo.gl/G3QdDQ6cRaVB8SN39",
    "website": "",
    "description": "Hoi An Ivy Hotel tọa lạc trên đường Đào Duy Từ, gần Chùa Cầu và một số điểm tham quan tại phố cổ Hội An. Phòng nghỉ được giới thiệu có máy điều hòa, phòng tắm riêng, tivi, minibar và Wi-Fi miễn phí; một số phòng có ban công hoặc tầm nhìn ra khu vực xung quanh. Các dịch vụ gồm lễ tân 24 giờ, dọn phòng hằng ngày, nhà hàng, quầy bar và xe đạp miễn phí. Vị trí này tạo điều kiện thuận lợi để du khách tham quan, thưởng thức ẩm thực và tìm hiểu văn hóa địa phương.",
    "nearby": "01 Đào Duy Từ, Cẩm Phô; cách Chùa Cầu ~300m",
    "rooms": "Standard, Superior, Deluxe",
    "amenities": "Wifi, Đặt vé/tour, Giặt ủi, Xe đạp",
    "food": "Khu vực ăn sáng nhỏ phục vụ khách lưu trú",
    "wellness": "",
    "events": "",
    "access": "Chỗ để xe máy, ô tô đỗ ven đường theo chỉ dẫn",
    "green": "",
    "guests": "Khách du lịch tiết kiệm, người đi du lịch một mình",
    "referencePrice": "350.000 – 700.000 VND",
    "referencePriceSource": "Agoda / Booking.com",
    "highlight": "Vị trí cực sát phố cổ, giá thành rẻ, chủ nhà nhiệt tình",
    "limitations": "Khách sạn nhỏ (Budget Hotel), không có hồ bơi hay spa",
    "excelImageNames": [
      "17_01_Mat_tien",
      "17_02_Phong_nghi",
      "17_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/3S05_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/3S05_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/3S05_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "HS01",
    "name": "Em's House Homestay",
    "group": "HS",
    "category": "Homestay",
    "type": "Homestay",
    "stars": "Không áp dụng",
    "address": "10/1 Đường Cao Hồng Lãnh, Hội An, Đà Nẵng, Việt Nam",
    "lat": 15.8777927706541,
    "lng": 108.32383288544,
    "maps": "https://maps.app.goo.gl/gC1BxUjiFVqDznd59",
    "website": "",
    "description": "Em's House Homestay là cơ sở lưu trú tọa lạc tại số 10/1 đường Cao Hồng Lãnh, Hội An. Homestay nằm gần khu vực phố cổ, thuận tiện cho du khách đi bộ tham quan chợ, các tuyến phố đèn lồng và chợ đêm Hội An. Cơ sở có phòng nghỉ tiện nghi, hồ bơi ngoài trời và khu vực sân vườn, đáp ứng nhu cầu nghỉ ngơi của khách du lịch. Không gian lưu trú tương đối yên tĩnh, kết hợp với vị trí thuận lợi giúp du khách dễ dàng khám phá các điểm tham quan và trải nghiệm cuộc sống địa phương. Đây là lựa chọn phù hợp cho khách du lịch cá nhân, cặp đôi và gia đình muốn tìm kiếm nơi lưu trú thoải mái gần phố cổ",
    "nearby": "Hẻm/kiệt Cẩm Phô; cách trung tâm Phố cổ ~500m",
    "rooms": "Double Room, Family Room",
    "amenities": "Bếp chung, Xe đạp miễn phí, Wifi, Phòng khách chung",
    "food": "Bữa sáng gia đình đơn giản (trứng, mỳ, cà phê)",
    "wellness": "",
    "events": "",
    "access": "Bãi để xe máy; ô tô gửi ở bãi ngoài phố",
    "green": "Không khí gia đình, dùng đồ dùng tái sử dụng trong sinh hoạt",
    "guests": "Khách trải nghiệm văn hóa địa phương, phượt thủ, ở dài ngày",
    "referencePrice": "300.000 – 600.000 VND",
    "referencePriceSource": "Booking.com",
    "highlight": "Thấu hiểu văn hóa địa phương, không gian ấm cúng, chủ nhà mến khách",
    "limitations": "Nằm trong ngõ, tiện nghi chuẩn homestay gia đình (không có thang máy)",
    "excelImageNames": [
      "18_01_Mat_tien",
      "18_02_Phong_nghi",
      "18_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/HS01_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/HS01_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/HS01_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "HS02",
    "name": "Hoi An Heart Lodge",
    "group": "HS",
    "category": "Homestay",
    "type": "Homestay",
    "stars": "Không áp dụng",
    "address": "621 621/7 Đ. Hai Bà Trưng, Hội An, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8810204094712,
    "lng": 108.32761728042,
    "maps": "https://maps.app.goo.gl/RMNa1C3wuVx4mGho6",
    "website": "",
    "description": "Hoi An Heart Lodge là cơ sở lưu trú theo mô hình homestay kết hợp phục vụ bữa sáng, tọa lạc tại số 621/7 đường Hai Bà Trưng, phường Cẩm Phô, Hội An. Cơ sở cung cấp phòng nghỉ có máy điều hòa, phòng tắm riêng, ban công và Wi-Fi miễn phí. Ngoài ra, homestay còn có hồ bơi trong nhà, sân tắm nắng, khu vực sinh hoạt chung, bếp dùng chung và xe đạp miễn phí. Bữa sáng có nhiều lựa chọn như món ăn Việt Nam, món Á, món Âu và thực đơn chay. Với vị trí cách phố cổ khoảng 5 phút đi bộ, không gian hiện đại và dịch vụ thân thiện, đây là lựa chọn phù hợp cho khách du lịch cá nhân, cặp đôi và du khách yêu thích sự tiện nghi",
    "nearby": "Đường Nguyễn Trường Tộ, Cẩm Phô; cách Phố cổ ~500m",
    "rooms": "Deluxe Double/Twin, Balcony Room",
    "amenities": "Wifi, Dịch vụ đưa đón, Cho thuê xe đạp/xe máy",
    "food": "Phục vụ ăn sáng tại sảnh/sân vườn",
    "wellness": "Hồ bơi ngoài trời cỡ nhỏ;  ;",
    "events": "",
    "access": "Chỗ để xe máy, ô tô đậu trước cổng",
    "green": "",
    "guests": "Cặp đôi, du khách nước ngoài, khách thích không gian nhỏ xinh",
    "referencePrice": "500.000 – 950.000 VND",
    "referencePriceSource": "Agoda",
    "highlight": "Đạt điểm đánh giá rất cao trên OTA nhờ dịch vụ chu đáo, vị trí tiện lợi",
    "limitations": "Quy mô nhỏ kiểu Lodge, số lượng phòng ít",
    "excelImageNames": [
      "19_01_Mat_tien",
      "19_02_Phong_nghi",
      "19_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/HS02_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/HS02_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/HS02_03.jpg",
    "reviewNote": "Sheet danh sách ghi 621/7 Hai Bà Trưng (ô địa chỉ bị lặp “621 621/7”), nhưng sheet chi tiết ghi Nguyễn Trường Tộ. Cần xác minh địa chỉ."
  },
  {
    "id": "HS03",
    "name": "Jolie Villa Hoi An",
    "group": "HS",
    "category": "Homestay",
    "type": "Homestay",
    "stars": "Không áp dụng",
    "address": "14 Lưu Trọng Lư, Hội An Tây, Đà Nẵng 560000, Việt Nam",
    "lat": 15.8909332709061,
    "lng": 108.322857233856,
    "maps": "https://maps.app.goo.gl/vesbE1RSSXpJhTQJA",
    "website": "",
    "description": "Jolie Villa Hoi An là cơ sở lưu trú tọa lạc tại số 14 đường Lưu Trọng Lư, khu vực Tân An, Hội An. Homestay cung cấp nhiều loại phòng nghỉ, bao gồm phòng gia đình có phòng tắm riêng, máy điều hòa và Wi-Fi miễn phí; một số phòng có ban công hoặc sân hiên nhìn ra vườn và hồ. Cơ sở có sân vườn, khu vực nghỉ ngơi chung, bếp dùng chung và xe đạp miễn phí để du khách khám phá các địa điểm tham quan xung quanh. Ngoài ra, homestay còn hỗ trợ dịch vụ đưa đón sân bay, đặt tour và phục vụ nhu cầu của khách trong thời gian lưu trú. Với không gian yên tĩnh và tiện nghi phù hợp, đây là lựa chọn dành cho gia đình, cặp đôi và nhóm khách muốn nghỉ ngơi kết hợp tham quan Hội An",
    "nearby": "Đường Lưu Trọng Lư, Tân An; cách Phố cổ ~1.8 km",
    "rooms": "Studio với bếp nhỏ, Deluxe Garden View",
    "amenities": "Lớp học nấu ăn, Xe đạp miễn phí, Sân vườn rộng",
    "food": "Phục vụ bữa sáng tươi ngon, tổ chức tiệc BBQ/lớp nấu ăn",
    "wellness": "Hồ bơi ngoài trời sân vườn;  ;",
    "events": "",
    "access": "Có bãi đỗ ô tô riêng",
    "green": "Trồng rau sạch trong khuôn viên, phân loại rác thân thiện môi trường",
    "guests": "Khách nước ngoài ở dài hạn, gia đình thích tự nấu nướng",
    "referencePrice": "500.000 – 1.000.000 VND",
    "referencePriceSource": "Booking.com / Agoda",
    "highlight": "Không gian gia đình cực kỳ ấm cúng, có các trải nghiệm ẩm thực tương tác cao",
    "limitations": "Hơi xa trung tâm phố cổ nếu chỉ muốn đi bộ",
    "excelImageNames": [
      "20_01_Mat_tien",
      "20_02_Phong_nghi",
      "20_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/HS03_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/HS03_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/HS03_03.jpg",
    "reviewNote": ""
  },
  {
    "id": "HS04",
    "name": "The View Homestay Hoi An",
    "group": "HS",
    "category": "Homestay",
    "type": "Homestay",
    "stars": "Không áp dụng",
    "address": "28/6 Trần Hưng Đạo, Hội An, Đà Nẵng, Việt Nam",
    "lat": 15.8804394905718,
    "lng": 108.329172891527,
    "maps": "https://maps.app.goo.gl/WpUYKkSD9k4QRidJ7",
    "website": "",
    "description": "The View Homestay Hoi An là cơ sở lưu trú tọa lạc tại số 28/6 đường Trần Hưng Đạo, phường Cẩm Phô, Hội An. Homestay có các phòng nghỉ được trang bị máy điều hòa, phòng tắm riêng, Wi-Fi miễn phí và một số tiện nghi như minibar, dụng cụ pha trà và cà phê. Một số phòng có tầm nhìn ra khu vườn, hồ bơi hoặc khu vực xung quanh. Cơ sở còn có hồ bơi trong nhà, sân vườn, sân hiên, khu vực sinh hoạt chung và dịch vụ lễ tân 24 giờ. Nhờ vị trí thuận tiện, du khách có thể đi bộ đến Bảo tàng Hội An, Chùa Cầu và các địa điểm ăn uống nổi tiếng trong khu phố cổ. Với không gian sạch sẽ, tiện nghi và vị trí trung tâm, homestay phù hợp với khách du lịch cá nhân, cặp đôi và gia đình",
    "nearby": "Khu vực Cẩm Nam (qua cầu Cẩm Nam); cách Phố cổ ~1 km",
    "rooms": "Standard Double, Room with Balcony/River View",
    "amenities": "Ban công hóng mát, Wifi, Mượn xe đạp",
    "food": "Phục vụ bữa sáng nhẹ (cà phê, bánh mì, trái cây)",
    "wellness": "",
    "events": "",
    "access": "Bãi xe máy, xe ô tô di chuyển vào được",
    "green": "",
    "guests": "Khách du lịch tiết kiệm, bạn trẻ thích ngắm cảnh",
    "referencePrice": "300.000 – 600.000 VND",
    "referencePriceSource": "Booking.com / Traveloka",
    "highlight": "Ban công view thoáng mát hướng ra sông/xóm làng, không gian yên bình",
    "limitations": "Tiện nghi cơ bản chuẩn homestay, không có thang máy",
    "excelImageNames": [
      "21_01_Mat_tien",
      "21_02_Phong_nghi",
      "21_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/HS04_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/HS04_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/HS04_03.jpg",
    "reviewNote": "Sheet danh sách ghi 28/6 Trần Hưng Đạo, nhưng sheet chi tiết mô tả Cẩm Nam/qua cầu Cẩm Nam. Cần đối chiếu phần khu vực lân cận."
  },
  {
    "id": "HS05",
    "name": "Phoenix Homestay Hoi An",
    "group": "HS",
    "category": "Homestay",
    "type": "Homestay",
    "stars": "Không áp dụng",
    "address": "06 Nguyễn Tuân, Hội An Tây, Đà Nẵng 84510, Việt Nam",
    "lat": 15.8886925652749,
    "lng": 108.322087988358,
    "maps": "https://maps.app.goo.gl/surCCSTtjDBwNUKM8",
    "website": "",
    "description": "Phoenix Homestay Hoi An là cơ sở lưu trú tọa lạc tại số 06 Nguyễn Tuân, Hội An. Homestay cung cấp nhiều loại phòng nghỉ, trong đó có phòng Superior diện tích 28 m² và phòng Deluxe diện tích 32 m². Các phòng được trang bị máy điều hòa, Wi-Fi, truyền hình, minibar, máy sấy tóc và phòng tắm có bồn tắm; một số phòng có ban công riêng. Giá phòng được công bố trên website tham khảo là 660.000 đồng/đêm đối với phòng Superior và 770.000 đồng/đêm đối với phòng Deluxe, kèm bữa sáng. Với hệ thống tiện nghi khá đầy đủ và nhiều lựa chọn phòng nghỉ, cơ sở phù hợp với khách du lịch cá nhân, cặp đôi và gia đình muốn khám phá Hội An (Phoenix Homestay Hoi An",
    "nearby": "Đường Phan Đình Phùng, Tân An; cách Phố cổ ~1.3 km",
    "rooms": "Superior King, Family Room",
    "amenities": "Dịch vụ giặt ủi, Cho thuê xe máy, Xe đạp miễn phí",
    "food": "Phục vụ bữa sáng theo yêu cầu",
    "wellness": "Hồ bơi ngoài trời nhỏ ngâm mình;  ;",
    "events": "",
    "access": "Có khu vực đỗ xe an toàn",
    "green": "",
    "guests": "Khách du lịch tự túc, nhóm bạn, gia đình nhỏ",
    "referencePrice": "400.000 – 800.000 VND",
    "referencePriceSource": "Agoda",
    "highlight": "Đô thị hóa vừa phải, phòng sạch sẽ ngăn nắp, giá siêu bình dân",
    "limitations": "Dịch vụ ở mức cơ bản, không có các dịch vụ giải trí cao cấp",
    "excelImageNames": [
      "22_01_Mat_tien",
      "22_02_Phong_nghi",
      "22_03_Tien_ich"
    ],
    "images": [
      {
        "src": "anh/HS05_01.jpg",
        "label": "Ảnh 1 · Mặt tiền / toàn cảnh",
        "credit": ""
      },
      {
        "src": "anh/HS05_02.jpg",
        "label": "Ảnh 2 · Phòng / tiện ích",
        "credit": ""
      }
    ],
    "optionalImagePath": "anh/HS05_03.jpg",
    "reviewNote": "Sheet danh sách ghi 06 Nguyễn Tuân, còn sheet chi tiết ghi đường Phan Đình Phùng. Cần đối chiếu thông tin địa điểm."
  }
];
