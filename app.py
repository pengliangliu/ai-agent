import os
import datetime
import json
import base64
import pandas as pd
import streamlit as st
from docx import Document as DocxDocument
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage, ToolMessage
from langchain_core.tools import tool
from langchain_core.documents import Document as LangchainDocument
from langchain_community.vectorstores import FAISS
from langchain_huggingface import HuggingFaceEmbeddings
from ddgs import DDGS

# ==========================================
# 0. 核心配置区
# ==========================================
DEEPSEEK_API_KEY = "sk-14da4b806d09469faaf06b14a2012673"  # ⚠️ 请填入你的真实 Key
os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"  # 强制 HuggingFace 使用国内镜像源

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(page_title="全能法规智能助理", page_icon="🤖", layout="wide")


# ==========================================
# 2. 轻量级身份验证系统
# ==========================================
def check_password():
    USER_CREDENTIALS = {"admin": "123456", "boss": "888888"}

    def password_entered():
        username = st.session_state["login_username"]
        password = st.session_state["login_password"]
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state["password_correct"] = True
            del st.session_state["login_password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 欢迎访问全能智能 Agent")
        st.text_input("用户名", key="login_username")
        st.text_input("密码", type="password", key="login_password")
        st.button("登录", on_click=password_entered, type="primary")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 欢迎访问全能智能 Agent")
        st.text_input("用户名", key="login_username")
        st.text_input("密码", type="password", key="login_password")
        st.button("登录", on_click=password_entered, type="primary")
        st.error("🚫 用户名或密码错误，请重试！")
        return False
    return True


# ==========================================
# 🌟 核心拦截器
# ==========================================
if check_password():

    # ==========================================
    # 3. 性能优化与工具定义
    # ==========================================
    @st.cache_resource
    def load_embedding_model():
        return HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")


    @tool
    def modify_word_document(original_text: str, revised_text: str, comment: str) -> str:
        """【文档修改工具】当用户同意修改方案，或要求“去修改Word文档”时调用此工具。"""
        if "current_file_path" not in st.session_state or not st.session_state.current_file_path:
            return "操作失败：当前没有加载任何文档。"
        source_path = st.session_state.current_file_path
        output_path = f"revised_{st.session_state.current_file_name}"
        try:
            doc = DocxDocument(source_path)
            modified_count = 0
            for paragraph in doc.paragraphs:
                if original_text in paragraph.text:
                    paragraph.text = paragraph.text.replace(original_text, revised_text)
                    for run in paragraph.runs:
                        if revised_text in run.text:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    run_comment = paragraph.add_run(f" [AI 批注: {comment}]")
                    run_comment.font.color.rgb = RGBColor(255, 0, 0)
                    modified_count += 1
            doc.save(output_path)
            if modified_count > 0:
                st.session_state.output_file_path = output_path
                return f"成功：已将修改写入文档并另存为 {output_path}。"
            return f"失败：未在原文档中找到原文片段 '{original_text}'。"
        except Exception as e:
            return f"修改文档时发生系统错误: {str(e)}"


    @tool
    def search_document_content(query: str) -> str:
        """【本地文档检索工具】当用户问起“文档里是怎么写的”时调用。"""
        if "vector_db" not in st.session_state or st.session_state.vector_db is None:
            return "操作失败：文档尚未被解析为向量数据库。"
        try:
            docs = st.session_state.vector_db.similarity_search(query, k=3)
            if not docs: return "检索完毕：没有找到相关内容。"
            context = "\n\n".join([f"原文段落 {i + 1}: {d.page_content}" for i, d in enumerate(docs)])
            return f"【本地检索结果】\n{context}"
        except Exception as e:
            return f"检索文档时发生错误: {str(e)}"


    @tool
    def search_latest_medical_regulations(query: str, time_limit: str = "m") -> str:
        """【互联网实时搜索工具】获取当前最新的法规、新闻。"""
        try:
            results = DDGS().text(query, max_results=3, timelimit=time_limit)
            if not results:
                return f"在 {time_limit} 范围内未能找到关于 '{query}' 的最新信息。"
            formatted_results = "\n\n".join(
                [f"标题: {res['title']}\n摘要: {str(res.get('body', ''))[:150]}...\n链接: {res['href']}" for res in
                 results])
            return f"【实时联网搜索结果】\n{formatted_results}"
        except Exception as e:
            return f"联网搜索发生异常: {str(e)}"


    @tool
    def generate_excel_matrix(json_data: str) -> str:
        """【Excel生成工具】生成NC整改矩阵物理文件。"""
        try:
            data = json.loads(json_data)
            df = pd.DataFrame(data)
            output_excel = "NC_Rectification_Matrix.xlsx"
            df.to_excel(output_excel, index=False)

            # 标记有新文件生成，由外部 Python 脚本负责注入 HTML 按钮
            st.session_state.excel_file_path = output_excel
            st.session_state.newly_generated_excel = output_excel

            return "成功：Excel文件已生成。请回复用户：'表格已为您生成，请点击下方的绿色按钮直接下载。'"
        except Exception as e:
            return f"生成 Excel 时发生错误: {str(e)}。请检查传入的 JSON 格式。"


    AVAILABLE_TOOLS = {
        "modify_word_document": modify_word_document,
        "search_document_content": search_document_content,
        "search_latest_medical_regulations": search_latest_medical_regulations,
        "generate_excel_matrix": generate_excel_matrix
    }


    def process_document_to_vector_db(file_path):
        doc = DocxDocument(file_path)
        paragraphs = [p.text for p in doc.paragraphs if len(p.text.strip()) > 5]
        if not paragraphs: return None
        docs = [LangchainDocument(page_content=text) for text in paragraphs]
        embeddings = load_embedding_model()
        vector_db = FAISS.from_documents(docs, embeddings)
        return vector_db


    # ==========================================
    # 4. 状态初始化与主界面布局
    # ==========================================
    st.title("🤖 医疗器械法规 Agent (沉浸式全能版)")

    if "current_file_path" not in st.session_state: st.session_state.current_file_path = None
    if "current_file_name" not in st.session_state: st.session_state.current_file_name = None
    if "output_file_path" not in st.session_state: st.session_state.output_file_path = None
    if "vector_db" not in st.session_state: st.session_state.vector_db = None
    if "excel_file_path" not in st.session_state: st.session_state.excel_file_path = None
    if "newly_generated_excel" not in st.session_state: st.session_state.newly_generated_excel = None

    # ==========================================
    # 5. 侧边栏
    # ==========================================
    with st.sidebar:
        st.header("⚙️ 系统状态")
        st.success("✅ 核心 AI 引擎已连接")
        st.markdown("---")
        st.header("📂 文档管理")
        uploaded_file = st.file_uploader("上传待修改的技术文档 (.docx)", type=["docx"])

        if uploaded_file is not None:
            if st.session_state.current_file_name != uploaded_file.name:
                save_path = f"temp_{uploaded_file.name}"
                with open(save_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state.current_file_path = save_path
                st.session_state.current_file_name = uploaded_file.name
                with st.spinner("正在解析文档并建立大脑索引..."):
                    st.session_state.vector_db = process_document_to_vector_db(save_path)
                st.success(f"已加载并解析: {uploaded_file.name}")
        else:
            st.session_state.current_file_path = None
            st.session_state.current_file_name = None
            st.session_state.output_file_path = None
            st.session_state.vector_db = None
            st.session_state.excel_file_path = None

        if st.session_state.output_file_path and os.path.exists(st.session_state.output_file_path):
            st.markdown("---")
            with open(st.session_state.output_file_path, "rb") as file:
                st.download_button("📥 下载修订版 Word", data=file, file_name=st.session_state.output_file_path,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   type="primary")

        st.markdown("---")
        if st.button("🚪 退出登录"):
            del st.session_state["password_correct"]
            st.rerun()

    if st.session_state.current_file_name:
        st.info(f"📄 **当前文档:** `{st.session_state.current_file_name}` | 🧠 **阅读状态:** 已载入记忆。")
    else:
        st.info("💡 **当前未加载文档。** 你可以把我当做全能助手进行日常闲聊，或上传文档进行深度分析。")

    # ==========================================
    # 6. Agent 聊天与核心调度逻辑
    # ==========================================
    real_today = datetime.datetime.now().strftime("%Y年%m月%d日")

    if "messages" not in st.session_state:
        system_prompt = f"""你是一个全能的AI智能助理，你的核心专长是资深医疗器械合规专家。
        【重要时间认知】：今天是真实的 {real_today}。
        你的工作模式：
        1. 【常规闲聊】：热情解答通用知识。
        2. 【联网模式】：调用 `search_latest_medical_regulations` 获取最新信息。
        3. 【文档处理】：调用本地文档搜索及 Word 修改工具。
        4. 【生成表格】：调用 `generate_excel_matrix` 工具生成 Excel 文件。
        """
        st.session_state.messages = [
            SystemMessage(content=system_prompt),
            AIMessage(content="你好！我是具备流式打字机和网页直出下载按钮能力的完全体 Agent。请问需要我做什么？")
        ]

    # 🚀 核心防抖机制：加入了 max_retries=3 和 timeout=60.0 防丢包报错
    llm = ChatOpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url="https://api.deepseek.com",
        model="deepseek-chat",
        temperature=0.3,
        streaming=True,
        max_retries=3,
        timeout=60.0
    )
    llm_with_tools = llm.bind_tools(list(AVAILABLE_TOOLS.values()))

    # 渲染历史对话 (开启 unsafe_allow_html 允许渲染内置按钮)
    for msg in st.session_state.messages:
        if isinstance(msg, SystemMessage) or isinstance(msg, ToolMessage): continue
        if isinstance(msg, AIMessage) and not msg.content: continue
        role = "user" if isinstance(msg, HumanMessage) else "assistant"
        with st.chat_message(role):
            st.markdown(msg.content, unsafe_allow_html=True)

    if user_input := st.chat_input("输入指令..."):

        # 记忆瘦身防卡顿 (保留系统提示词 + 最近10条消息)
        if len(st.session_state.messages) > 11:
            st.session_state.messages = [st.session_state.messages[0]] + st.session_state.messages[-10:]

        with st.chat_message("user"):
            st.markdown(user_input, unsafe_allow_html=True)
        st.session_state.messages.append(HumanMessage(content=user_input))

        with st.chat_message("assistant"):
            max_loops = 8
            current_loop = 0

            while current_loop < max_loops:
                current_loop += 1
                message_placeholder = st.empty()
                ai_msg_chunk = None
                status_text = "AI 正在思考..." if current_loop == 1 else "AI 正在分析执行结果..."

                with st.spinner(status_text):
                    # 流式渲染打字机效果
                    for chunk in llm_with_tools.stream(st.session_state.messages):
                        if ai_msg_chunk is None:
                            ai_msg_chunk = chunk
                        else:
                            ai_msg_chunk = ai_msg_chunk + chunk

                        if ai_msg_chunk.content:
                            message_placeholder.markdown(ai_msg_chunk.content + " ▌", unsafe_allow_html=True)

                # 去掉光标
                if ai_msg_chunk.content:
                    message_placeholder.markdown(ai_msg_chunk.content, unsafe_allow_html=True)

                st.session_state.messages.append(ai_msg_chunk)

                if not ai_msg_chunk.tool_calls:
                    break

                # 后台静默执行工具
                for tool_call in ai_msg_chunk.tool_calls:
                    tool_name = tool_call["name"]
                    tool_args = tool_call["args"]
                    tool_func = AVAILABLE_TOOLS.get(tool_name)

                    try:
                        if tool_func:
                            result_msg = tool_func.invoke(tool_args)
                        else:
                            result_msg = f"系统错误：找不到名为 {tool_name} 的工具。"
                    except Exception as e:
                        result_msg = f"工具执行异常: {str(e)}"

                    tool_message = ToolMessage(content=str(result_msg), tool_call_id=tool_call["id"])
                    st.session_state.messages.append(tool_message)

            if current_loop >= max_loops and ai_msg_chunk.tool_calls:
                st.warning("⚠️ 思考及搜索轮数达到上限，已强制中断推导。")

            # 🌟 核心拦截器：Python 强行介入贴上 Excel 下载按钮
            if st.session_state.get("newly_generated_excel"):
                file_path = st.session_state.newly_generated_excel
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode()

                    # 构造原生的 HTML 绿色按钮
                    html_link = f'<div style="margin-top: 15px;"><a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_path}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #00A67E; color: white; text-align: center; text-decoration: none; border-radius: 6px; font-weight: bold;">📥 点击这里直接下载 Excel 矩阵表</a></div>'

                    # 渲染在气泡底部
                    st.markdown(html_link, unsafe_allow_html=True)

                    # 悄悄拼接进记忆，保证刷新不丢按钮
                    if len(st.session_state.messages) > 0 and isinstance(st.session_state.messages[-1], AIMessage):
                        st.session_state.messages[-1].content += "\n\n" + html_link

                # 清除标记，防止下一个对话乱弹按钮
                st.session_state.newly_generated_excel = None