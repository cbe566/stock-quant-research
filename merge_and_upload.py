"""
合併 4 個市場的篩選結果 → 生成完整報告 → 上傳到 D1
用於 GitHub Actions matrix job 的最後合併步驟
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

def merge_market_results():
    """合併各市場 JSON 結果為完整報告"""
    artifacts_dir = Path("market_results")
    if not artifacts_dir.exists():
        print("❌ market_results/ 目錄不存在")
        sys.exit(1)

    # 讀取所有市場結果
    all_markets = {}
    all_errors = []
    date_str = None

    for json_file in sorted(artifacts_dir.glob("screening_*.json")):
        print(f"  📄 讀取 {json_file.name}")
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        market = data.get("market", "unknown")
        all_markets[market] = data.get("results", [])
        all_errors.extend(data.get("errors", []))
        if not date_str:
            date_str = data.get("date")

    if not all_markets:
        print("❌ 沒有找到任何市場結果")
        sys.exit(1)

    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    print(f"\n✅ 合併了 {len(all_markets)} 個市場的結果")
    for market, results in all_markets.items():
        print(f"   {market}: {len(results)} 支股票")

    # 儲存合併後的完整 JSON
    report_dir = Path("daily_reports")
    report_dir.mkdir(exist_ok=True)

    date_compact = date_str.replace("-", "")
    json_path = report_dir / f"screening_data_{date_compact}.json"
    merged_data = {
        "date": date_str,
        "generated_at": datetime.now().isoformat(),
        "markets": all_markets,
        "errors": all_errors,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(merged_data, f, ensure_ascii=False, indent=2, default=str)
    print(f"📊 合併數據已存到 {json_path}")

    # 生成 Markdown 報告（用 DailyScreener 的報告邏輯）
    try:
        sys.path.insert(0, os.path.dirname(__file__))
        from daily_screening import DailyScreener
        screener = DailyScreener()
        screener.results = all_markets
        screener.errors = all_errors
        md_path, report = screener.save_report()
        print(f"📄 Markdown 報告已生成：{md_path}")
    except Exception as e:
        print(f"⚠️ 報告生成失敗：{e}")
        # 至少保存 JSON 就好
        import traceback
        traceback.print_exc()

    return str(json_path)


def upload_to_d1(json_path=None):
    """上傳結果到 Cloudflare D1"""
    try:
        # 嘗試用現有的 upload_to_d1.py
        sys.path.insert(0, os.path.dirname(__file__))
        from upload_to_d1 import upload_screening_data
        success = upload_screening_data(json_path)
        if success:
            print("✅ D1 上傳成功")
        else:
            print("⚠️ D1 上傳回傳失敗")
    except Exception as e:
        print(f"⚠️ D1 上傳失敗（不影響報告）：{e}")


def generate_pdf(json_path):
    """生成 PDF 詳細報告"""
    try:
        from generate_report_pdf import DailyReportPDF
        gen = DailyReportPDF(json_path)
        pdf_path = gen.generate()
        print(f"📄 PDF 報告已生成：{pdf_path}")
        return pdf_path
    except ImportError as e:
        print(f"⚠️ PDF 生成跳過（缺少 reportlab）：{e}")
        return None
    except Exception as e:
        print(f"⚠️ PDF 生成失敗：{e}")
        import traceback
        traceback.print_exc()
        return None


def send_email(json_path, pdf_path=None):
    """透過 Gmail OAuth2 API 發送報告郵件"""
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    gmail_user = os.environ.get("GMAIL_USER")
    client_id = os.environ.get("GMAIL_CLIENT_ID")
    client_secret = os.environ.get("GMAIL_CLIENT_SECRET")
    refresh_token = os.environ.get("GMAIL_REFRESH_TOKEN")
    email_to = os.environ.get("EMAIL_TO", gmail_user)

    if not all([gmail_user, client_id, client_secret, refresh_token]):
        print("⚠️ 未設定 Gmail OAuth2 憑證，跳過郵件發送")
        return False

    try:
        from send_daily_email import generate_email_html, generate_email_subject, generate_email_plain

        subject = generate_email_subject(json_path)
        html_body = generate_email_html(json_path)
        plain_body = generate_email_plain(json_path)

        # 組裝 MIME 郵件
        msg = MIMEMultipart("mixed")
        msg["From"] = f"何宣逸 <{gmail_user}>"
        msg["To"] = email_to
        msg["Subject"] = subject

        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain_body, "plain", "utf-8"))
        alt.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alt)

        # 附件 PDF
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf_att = MIMEApplication(f.read(), _subtype="pdf")
                pdf_filename = os.path.basename(pdf_path)
                pdf_att.add_header("Content-Disposition", "attachment", filename=pdf_filename)
                msg.attach(pdf_att)
            print(f"📎 附件：{pdf_filename}")

        # 用 refresh_token 換取 access_token
        token_resp = requests.post("https://oauth2.googleapis.com/token", data={
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        })
        if token_resp.status_code != 200:
            print(f"⚠️ OAuth2 token 刷新失敗：{token_resp.text[:200]}")
            return False

        access_token = token_resp.json()["access_token"]

        # 透過 Gmail API 發送
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        send_resp = requests.post(
            "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            json={"raw": raw},
            timeout=30,
        )

        if send_resp.status_code == 200:
            print(f"📧 郵件已發送到 {email_to}")
            return True
        else:
            print(f"⚠️ 郵件發送失敗：HTTP {send_resp.status_code} - {send_resp.text[:200]}")
            return False

    except Exception as e:
        print(f"⚠️ 郵件發送失敗：{e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    json_path = merge_market_results()
    upload_to_d1(json_path)
    pdf_path = generate_pdf(json_path)
    send_email(json_path, pdf_path)
