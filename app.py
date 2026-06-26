import streamlit as st
from anthropic import Anthropic

st.set_page_config(page_title="Gerador de Ebooks", page_icon="📚")

st.title("📚 Gerador de Ebooks Premium")
st.markdown("Crie ebooks profissionais com IA")

# Inputs
topic = st.text_input("📖 Tema do Ebook", "Maternidade real")
audience = st.text_input("👥 Público-alvo", "mães e pais")

if st.button("🚀 Gerar Ebook", use_container_width=True):
    try:
        client = Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
        
        st.info("⏳ Gerando conteúdo com Claude...")
        
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=2000,
            messages=[{
                "role": "user",
                "content": f"""Crie um ebook sobre "{topic}" para {audience}.

Estrutura JSON:
{{
  "title": "Título",
  "subtitle": "Subtítulo",
  "chapters": [
    {{"number": 1, "title": "Cap 1", "content": "Conteúdo..."}},
    {{"number": 2, "title": "Cap 2", "content": "Conteúdo..."}},
    {{"number": 3, "title": "Cap 3", "content": "Conteúdo..."}}
  ]
}}"""
            }]
        )
        
        text = response.content[0].text
        st.success("✅ Ebook gerado!")
        st.write(text)
        
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

st.markdown("---")
st.markdown("**Luciana Britto | L&B Marketing**")
