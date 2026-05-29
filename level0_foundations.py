"""
Level 0 — Foundations
=====================
เป้าหมาย:
  1) โหลด ANTHROPIC_API_KEY จาก .env
  2) init chat model ด้วย API v1 (init_chat_model)
  3) invoke LLM ครั้งแรก
  4) เข้าใจ message types: System / Human / AI

รัน:  ./venv/bin/python level0_foundations.py
"""

from dotenv import load_dotenv
from langchain.chat_models import init_chat_model
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage

# 1) โหลด environment variables จากไฟล์ .env (อ่าน ANTHROPIC_API_KEY เข้ามา)
load_dotenv()

# 2) สร้าง chat model แบบ provider-agnostic (API v1)
#    - รูปแบบ "provider:model_name" อ่านง่ายและเปลี่ยน provider ได้ทันที
#    - temperature=0 = ตอบนิ่ง/เดาน้อย เหมาะกับการเรียน (reproducible)
model = init_chat_model(
    "anthropic:claude-sonnet-4-6",
    temperature=0,
)

# 3) invoke แบบง่ายสุด: ส่ง string เข้าไปตรงๆ
#    LangChain จะแปลง string ให้เป็น HumanMessage ให้อัตโนมัติ
print("=" * 60)
print("ตัวอย่างที่ 1: invoke ด้วย string ธรรมดา")
print("=" * 60)
resp = model.invoke("อธิบาย 'LangGraph คืออะไร' สั้นๆ ใน 2 ประโยค")
print(resp.content)

# 4) invoke ด้วย list ของ messages (วิธีที่ใช้จริงตลอดคอร์ส)
#    SystemMessage = ตั้งบุคลิก/กฎ | HumanMessage = คำถามจากผู้ใช้
print("\n" + "=" * 60)
print("ตัวอย่างที่ 2: invoke ด้วย list[messages] + SystemMessage")
print("=" * 60)
messages = [
    SystemMessage(content="คุณคือ mentor ด้าน AI ที่ตอบกระชับ เป็นภาษาไทย และยกตัวอย่างเสมอ"),
    HumanMessage(content="tool calling คืออะไร?"),
]
resp2 = model.invoke(messages)
print(resp2.content)

# 5) สำรวจ "ของที่ติดมากับ AIMessage" — สำคัญมากสำหรับ level ถัดๆ ไป
print("\n" + "=" * 60)
print("ตัวอย่างที่ 3: ส่อง metadata ของ AIMessage")
print("=" * 60)
print("type ของ response :", type(resp2).__name__)          # AIMessage
print("model ที่ใช้ตอบ   :", resp2.response_metadata.get("model"))
print("token ที่ใช้ไป    :", resp2.usage_metadata)            # input/output/total tokens
print("tool_calls       :", resp2.tool_calls)                # ตอนนี้ว่าง [] — จะมีค่าใน Level 1
