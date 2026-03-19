import sys
import base64
try:
    from openai import OpenAI
except ImportError:
    import subprocess
    print("正在安装OpenAI...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "openai"])
    from openai import OpenAI

_client = None
DEFAULT_MODEL = "qwen-vl-plus"

def setup_client(api_key: str, base_url: str):
    """
    配置 OpenAI 客户端
    :param api_key: API Key
    :param base_url: Base URL
    """
    global _client
    _client = OpenAI(
        api_key=api_key,
        base_url=base_url
    )

def encode_image(image_path: str) -> str:
    """
    将本地图片转为 base64 字符串
    :param image_path: 图片路径
    :return: base64 编码字符串
    """
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

DEFAULT_SYSTEM_PROMPT_BASE = (
    "你是一个招聘助理，负责分析简历图片。请执行以下步骤：\n"
    "1. 使用 OCR 识别简历上的所有文字，特别关注'备注'区域的内容。\n"
    "2. 逐条分析备注记录，判断是否建议再次联系。\n"
    "3. **核心判断规则**（严格执行）：\n"
    "   - **拒绝条件**：仔细检查每一条备注。如果发现某条备注的主要内容仅仅是 '1'、'2' 或 'p'（不区分大小写），则判定为”不需要联系“。\n"
    "   - **常见形式**：\n"
    "     - 单独一行的 '1'。\n"
    "     - **特别注意**：'仅自己可见' 是备注框的默认提示或属性，**不能**作为判断依据。必须是用户输入的 '1' 或 '2' 或 'p' 才是拒绝理由。\n"
    "     - 如果 OCR 识别结果是 '1 仅自己可见'，可以认为是拒绝（因为包含 '1'）。\n"
    "     - 但如果 OCR 识别结果**仅仅**是 '仅自己可见'，或者 '备注' 标题，则**完全忽略**。\n"
    "   - **排除干扰**：\n"
    "     - 如果 '1' 出现在手机号中（如 138...），**不**算拒绝。\n"
    "     - 如果 '1' 是长句子的一部分（如 '工作经验1年'），**不**算拒绝。\n"
    "     - 忽略 '备注(3)' 标题中的数字。\n"
    "   - 如果没有发现符合上述拒绝条件的备注，回答“需要联系”。\n"
)

DEFAULT_OUTPUT_FORMAT = (
    "4. 你的回答格式必须是：判定结果|判定依据。\n"
    "   - 判定结果只能是”需要联系“或者”不需要联系“。\n"
    "   - 判定依据简要说明原因。\n"
    "   - 示例：不需要联系|发现备注内容为 '1' (或 '1 仅自己可见')\n"
    "   - 示例：需要联系|未发现拒绝类备注（手机号不视为拒绝）\n"
)

DEFAULT_SYSTEM_PROMPT = DEFAULT_SYSTEM_PROMPT_BASE + DEFAULT_OUTPUT_FORMAT

def evaluate_resume_from_image(image_path: str, user_text: str, system_prompt: str = None, model: str = None) -> str:
    """
    分析简历图片，提取备注，并判断是否需要联系
    
    :param image_path: 图片路径
    :param user_text: 用户指令文本
    :param system_prompt: 系统提示词（可选，如果提供，将作为补充规则追加到默认规则之后）
    :param model: 使用的 AI 模型（可选，默认为 qwen-vl-plus）
    :return: 分析结果文本或 JSON 字符串
    """
    global _client
    if _client is None:
        return "错误|AI 客户端未配置，请先在设置中配置 API Key"

    # Use default model if not provided
    if not model or not model.strip():
        model = DEFAULT_MODEL

    try:
        base64_img = encode_image(image_path)
    except Exception as e:
         return f"错误|读取图片失败: {str(e)}"

    # Construct System Prompt
    # Logic: Always start with BASE rules.
    # If user provided a custom prompt, append it as "Additional User Rules".
    # Always end with OUTPUT FORMAT to ensure correct parsing.
    
    final_prompt = DEFAULT_SYSTEM_PROMPT_BASE
    
    if system_prompt and system_prompt.strip():
        final_prompt += f"\n5. **用户额外补充规则**（优先级高于默认规则）：\n{system_prompt}\n"
    
    final_prompt += DEFAULT_OUTPUT_FORMAT

    # Ensure user_text is not empty to trigger the model properly
    if not user_text or not user_text.strip():
        user_text = "请分析图片中的备注内容，并严格按照指定格式输出结果。"

    messages = [
        {"role": "system", "content": final_prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": user_text},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
            ]
        }
    ]

    try:
        response = _client.chat.completions.create(
            model=model,
            messages=messages,
            max_tokens=512,
            temperature=0.1  # 降低随机性，确保稳定输出 JSON
        )
        return response.choices[0].message.content
    except Exception as e:
        error_msg = str(e)
        if "Arrearage" in error_msg:
            return f"错误|AI账户欠费或额度不足 (Arrearage)，请充值后重试。"
        return f"错误|AI请求失败: {error_msg}"

if __name__ == "__main__":
    pass
