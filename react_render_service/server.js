const express = require('express');
const cors = require('cors');
const React = require('react');
const { Presentation, Slide, Text, render } = require('react-pptx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

app.get('/', (req, res) => {
    res.send('🚀 Trạm Render React-PPTX đang hoạt động ngon lành!');
});

app.post('/api/render-pptx', async (req, res) => {
    try {
        const slideData = req.body;
        console.log("📥 Đã nhận yêu cầu từ Python! Số slide AI tạo ra:", slideData.slides?.length || 0);

        const MyPresentation = () => (
            <Presentation>
                {/* 1. SLIDE BÌA (Trang mở đầu) */}
                <Slide>
                    {/* Sử dụng số thực (Inch) thay vì phần trăm */}
                    <Text style={{ 
                        x: 1, y: 2, w: 8, h: 1, 
                        fontSize: 44, color: '#003366', bold: true, align: 'center' 
                    }}>
                        BÁO CÁO ĐA DỤNG (UNIVERSAL)
                    </Text>
                    <Text style={{ 
                        x: 1, y: 3, w: 8, h: 1, 
                        fontSize: 22, color: '#009975', align: 'center' 
                    }}>
                        Tạo tự động bởi AI Core & React Engine
                    </Text>
                </Slide>

                {/* 2. CÁC SLIDE NỘI DUNG TỪ AI */}
                {slideData.slides && slideData.slides.map((s, index) => (
                    <Slide key={index}>
                        
                        {/* Tiêu đề (Cách lề trái 0.5 inch, lề trên 0.4 inch) */}
                        <Text style={{ 
                            x: 0.5, y: 0.4, w: 9, h: 0.8, 
                            fontSize: 32, color: '#003366', bold: true 
                        }}>
                            {s.title || "Slide chưa có tiêu đề"}
                        </Text>

                        {/* Các Bullet Points */}
                        {s.bullets && s.bullets.map((b, i) => (
                            <Text key={i} style={{
                                x: 0.8, 
                                y: 1.5 + (i * 0.8), // Dòng đầu ở tọa độ y=1.5, các dòng sau cách nhau 0.8 inch
                                w: 8.5, h: 0.6,
                                fontSize: 18, color: '#334155'
                            }}>
                                {`• ${b}`}
                            </Text>
                        ))}

                        {/* Takeaway Banner (Cố định ở mép dưới cùng, y=4.8 inch) */}
                        {s.takeaway && (
                            <Text style={{
                                x: 0.5, y: 4.8, w: 9, h: 0.5,
                                fontSize: 14, color: '#d4af37', bold: true, align: 'center'
                            }}>
                                {`TAKEAWAY: ${s.takeaway}`}
                            </Text>
                        )}
                        
                    </Slide>
                ))}
            </Presentation>
        );

        // Dịch code React thành file PowerPoint Buffer
        const buffer = await render(<MyPresentation />);

        // Gửi trả file PPTX về cho Python
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.setHeader('Content-Disposition', 'attachment; filename="Universal_Report.pptx"');
        res.send(Buffer.from(buffer)); 
        
        console.log("✅ Đã xuất file PPTX chứa chữ thành công!");

    } catch (error) {
        console.error("❌ Lỗi Render:", error);
        res.status(500).json({ error: error.message });
    }
});

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`🚀 Trạm Render (Bản Tọa độ Inch Chuẩn) đang chạy tại http://localhost:${PORT}`);
});