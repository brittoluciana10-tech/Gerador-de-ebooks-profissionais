import streamlit as st
from anthropic import Anthropic
import json

st.set_page_config(page_title="Gerador de Ebooks", page_icon="📚")

st.title("📚 Gerador de Ebooks")
st.markdown("Crie ebooks com IA em segundos")

# Inputs
col1, col2 = st.columns(2)
with col1:
    topic = st.text_input("Tema", "Maternidade real")
with col2:
    audience = st.text_input("Público-alvo", "mães")

if st.button("🚀 Gerar", use_container_width=True):
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY")
        if not api_key:
            st.error("❌ Configure ANTHROPIC_API_KEY nos Secrets")
        else:
            client = Anthropic(api_key=api_key)
            
            with st.spinner("⏳ Gerando..."):
                response = client.messages.create(
                    model="claude-opus-4-6",
                    max_tokens=1500,
                    messages=[{
                        "role": "user",
                        "content": f"Crie 3 capítulos sobre '{topic}' para {audience}. Estrutura: título, 2 seções de 200 palavras cada."
                    }]
                )
                
                content = response.content[0].text
                st.success("✅ Ebook gerado!")
                st.write(content)
                
                # Botão para copiar
                st.code(content)
                
    except Exception as e:
        st.error(f"❌ {str(e)}")

st.markdown("---")
st.markdown("**Luciana Britto | L&B Marketing**")
