import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Gerador de Ebooks", page_icon="📚")

st.title("📚 Gerador de Ebooks")
st.markdown("Crie ebooks com IA Google (GRÁTIS!)")

topic = st.text_input("📖 Tema do Ebook", "Maternidade real")
audience = st.text_input("👥 Público-alvo", "mães")

if st.button("🚀 Gerar Ebook", use_container_width=True):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        model = genai.GenerativeModel("gemini-pro")
        
        with st.spinner("⏳ Gerando com Google Gemini..."):
            response = model.generate_content(
                f"Crie um ebook sobre '{topic}' para '{audience}'. "
                f"Estrutura: 3 capítulos, cada um com título e 200 palavras."
            )
            
            st.success("✅ Ebook gerado com sucesso!")
            st.write(response.text)
            
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

st.markdown("---")
st.markdown("**Luciana Britto | L&B Marketing — Estratégias de Valor**")
