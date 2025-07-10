import streamlit as st
import pandas as pd
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import base64

# 日本語フォントの設定
try:
    # reportlab-japanese-fontsがインストールされている場合
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
    JAPANESE_FONT = 'HeiseiMin-W3'
except ImportError:
    # フォールバック: システムフォントを使用
    try:
        # Windows の日本語フォント
        pdfmetrics.registerFont(TTFont('YuGothic', 'C:/Windows/Fonts/yu Gothic.ttc'))
        JAPANESE_FONT = 'YuGothic'
    except:
        # デフォルトフォント
        JAPANESE_FONT = 'Helvetica'

# チェックシートのデータ（カテゴリと項目）
data = {
    "1. 会社情報と体制": [
        "会社概要・所在地・人員構成",
        "組織図（最新版）と役割の明確化",
        "主要な受託業務と実績の記録"
    ],
    "2. 品質保証と体制": [
        "QA/QCの体制と人員",
        "SOPに基づく業務履行、管理、改訂の体制",
        "社内QA/QCのレビュー結果とCAPA（是正措置）管理",
        "SOPの管理・改訂履歴の確認"
    ],
    "3. データ・文書管理": [
        "データのバックアップ体制（頻度、手段）",
        "電子データの保管場所・セキュリティ",
        "文書保管規定・旧版管理の有無"
    ],
    "4. 教育・訓練": [
        "教育研修SOPの整備状況",
        "教育訓練の記録・更新状況",
        "専門的スキル・資格の保有状況"
    ],
    "5. システムとセキュリティ": [
        "施設の入退室管理（物理的セキュリティ）",
        "システムバリデーション（CSV）の有無",
        "クラウドやNASの安全性とログ管理"
    ],
    "6. プロジェクト管理": [
        "プロジェクト指名書・責任者の明確化",
        "業務手順の一貫性と記録の整備"
    ]
}

st.set_page_config(page_title="外部信頼性保証チェック＆レビューシート（内部監査相当）", layout="wide")
st.title("外部信頼性保証チェック＆レビューシート（内部監査相当）")
st.markdown("""
※ 本レビューは、SOP001 Annex-001-Aに基づき、当社が実施する外部QAレビューを記録するものであり、内部監査に準ずる品質保証活動として位置づけられます。  
※ レビュー実施者は、信頼性保証に関する知識・経験を有する適格な外部QA担当者によって実施されています。
""")

st.markdown("""
株式会社リジット  
コード開発者：代表および信頼性保証責任者  
山本修司
""")

auditor = st.text_input("レビュー実施者（外部QA)")
audit_date = st.date_input("実施日", value=date.today())

records = []

for section, items in data.items():
    st.header(section)
    for item in items:
        col1, col2 = st.columns([2, 3])
        col1.markdown(f"**項目：{item}**")
        status = col1.radio("対応状況", ["未確認", "確認済", "要修正"], key=f"status_{item}")
        comment = col2.text_input("コメント", key=f"comment_{item}")
        records.append({
            "カテゴリ": section,
            "項目": item,
            "対応状況": status,
            "コメント": comment
        })

def generate_pdf_report(records, auditor, audit_date, special_notes, evaluations):
    """PDFレポートを生成する関数"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # 日本語フォントを使用したスタイル設定
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue,
        fontName=JAPANESE_FONT
    )
    
    # 日本語フォントを使用した通常スタイル
    normal_style = ParagraphStyle(
        'JapaneseNormal',
        parent=styles['Normal'],
        fontName=JAPANESE_FONT,
        fontSize=10
    )
    
    # 日本語フォントを使用した見出しスタイル
    heading_style = ParagraphStyle(
        'JapaneseHeading',
        parent=styles['Heading2'],
        fontName=JAPANESE_FONT,
        fontSize=14,
        spaceAfter=10
    )
    
    # ヘッダー情報
    story.append(Paragraph("Imaging CRO 内部監査チェックシート", title_style))
    story.append(Spacer(1, 20))
    
    # 会社情報
    company_info = f"""
    <b>株式会社リジット</b><br/>
    コード開発者：代表および信頼性保証責任者<br/>
    山本修司
    """
    story.append(Paragraph(company_info, normal_style))
    story.append(Spacer(1, 20))
    
    # 監査情報
    audit_info = f"""
    <b>監査者：</b>{auditor}<br/>
    <b>監査日：</b>{audit_date}
    """
    story.append(Paragraph(audit_info, normal_style))
    story.append(Spacer(1, 30))
    
    # 各カテゴリの結果
    current_category = ""
    for record in records:
        if record['カテゴリ'] != current_category:
            current_category = record['カテゴリ']
            story.append(Paragraph(f"<b>{current_category}</b>", heading_style))
            story.append(Spacer(1, 10))
        
        # 項目と結果
        item_text = f"""
        <b>項目：</b>{record['項目']}<br/>
        <b>対応状況：</b>{record['対応状況']}<br/>
        <b>コメント：</b>{record['コメント']}
        """
        story.append(Paragraph(item_text, normal_style))
        story.append(Spacer(1, 10))
    
    # 特記事項
    story.append(Paragraph("<b>7. 特記事項</b>", heading_style))
    story.append(Paragraph(special_notes, normal_style))
    story.append(Spacer(1, 20))
    
    # 評価結果
    story.append(Paragraph("<b>監査結果の評価（顧問記入）</b>", heading_style))
    eval_data = [
        ['評価項目', '評価'],
        ['総合的な体制整備', evaluations['総合的な体制整備']],
        ['SOP運用の実効性', evaluations['SOP運用の実効性']],
        ['データ管理とセキュリティ', evaluations['データ管理とセキュリティ']],
        ['継続的改善の姿勢', evaluations['継続的改善の姿勢']]
    ]
    
    eval_table = Table(eval_data, colWidths=[3*inch, 1*inch])
    eval_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), JAPANESE_FONT),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('FONTNAME', (0, 1), (-1, -1), JAPANESE_FONT),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(eval_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# 7. 特記事項
st.header("7. 特記事項")
st.text_area("内部監査で確認された特記事項", height=100, key="special_notes")

# 監査結果の評価
st.subheader("監査結果の評価（顧問記入）")
eval_options = ["優", "良", "可", "不可"]
eval_1 = st.selectbox("総合的な体制整備", eval_options)
eval_2 = st.selectbox("SOP運用の実効性", eval_options)
eval_3 = st.selectbox("データ管理とセキュリティ", eval_options)
eval_4 = st.selectbox("継続的改善の姿勢", eval_options)

if st.button("保存する"):
    df = pd.DataFrame(records)
    df["特記事項"] = st.session_state["special_notes"]
    df["総合的な体制整備"] = eval_1
    df["SOP運用の実効性"] = eval_2
    df["データ管理とセキュリティ"] = eval_3
    df["継続的改善の姿勢"] = eval_4
    
    # ファイル名のベース部分
    base_filename = f"audit_{audit_date}_{auditor.replace(' ', '_')}"
    
    # CSVファイルの保存
    csv_filename = f"{base_filename}.csv"
    df.to_csv(csv_filename, index=False)
    
    # PDFファイルの生成
    evaluations = {
        "総合的な体制整備": eval_1,
        "SOP運用の実効性": eval_2,
        "データ管理とセキュリティ": eval_3,
        "継続的改善の姿勢": eval_4
    }
    
    pdf_buffer = generate_pdf_report(records, auditor, audit_date, st.session_state["special_notes"], evaluations)
    pdf_filename = f"{base_filename}.pdf"
    
    st.success(f"保存しました: {csv_filename}, {pdf_filename}")
    
    # ダウンロードボタン
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "CSVをダウンロード", 
            data=df.to_csv(index=False), 
            file_name=csv_filename, 
            mime="text/csv"
        )
    with col2:
        st.download_button(
            "PDFをダウンロード", 
            data=pdf_buffer.getvalue(), 
            file_name=pdf_filename, 
            mime="application/pdf"
        )
