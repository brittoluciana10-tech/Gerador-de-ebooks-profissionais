import streamlit as st
from openai import OpenAI

st.set_page_config(page_title="Gerador de Ebooks", page_icon="📚")

st.title("📚 Gerador de Ebooks")

topic = st.text_input("Tema do Ebook")
audience = st.text_input("Público-alvo")

if st.button("🚀 Gerar Ebook"):
    if not topic or not audience:
        st.warning("Preencha tema e público-alvo!")
    else:
        try:
            api_key = st.secrets["OPENAI_API_KEY"]
            client = OpenAI(api_key=api_key)
            
            with st.spinner("Gerando..."):
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    max_tokens=1000,
                    messages=[{
                        "role": "user",
                        "content": f"Crie um ebook sobre '{topic}' para '{audience}'. 3 capítulos com 150 palavras cada."
                    }]
                )
                
                st.success("✅ Pronto!")
                st.write(response.choices[0].message.content)
                
        except Exception as e:
            st.error(f"Erro: {str(e)}")
