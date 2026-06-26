import streamlit as st
import google.generativeai as genai
from fpdf import FPDF
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Gerador de Ebooks Premium", page_icon="📚", layout="wide")

# CSS customizado
st.markdown("""
    <style>
    .main { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
    </style>
""", unsafe_allow_html=True)

st.title("📚 Gerador de Ebooks Premium")
st.markdown("Crie ebooks profissionais com IA Google (GRÁTIS!) + PDF + Página de Venda")

# Temas disponíveis
TEMAS = {
    "🤰 Maternidade Real": "Maternidade real, mães, experiências honestas",
    "👶 Sono do Bebê": "Sono infantil, técnicas de dormir, rotina",
    "💪 Pós-parto": "Recuperação pós-parto, saúde, bem-estar",
    "📚 Parenting": "Educação infantil, desenvolvimento, parentalidade",
    "💰 Renda Variável": "Renda extra, freelance, negócios digitais",
    "🧘 Bem-estar": "Mindfulness, meditação, saúde mental",
    "💼 Autônomo": "Ser autônomo, gestão, financeiro",
}

col1, col2 = st.columns(2)

with col1:
    tema_nome = st.selectbox("📖 Escolha o Tema", list(TEMAS.keys()))
    tema_desc = TEMAS[tema_nome]

with col2:
    audience = st.text_input("👥 Público-alvo", "mães")

# Opções avançadas
with st.expander("⚙️ Opções Avançadas"):
    num_chapters = st.slider("Número de capítulos", 2, 5, 3)
    style = st.selectbox("Estilo", ["Profissional", "Descontraído", "Científico"])

if st.button("🚀 Gerar Ebook Premium", use_container_width=True):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        # Descobrir modelo disponível
        models = [m.name for m in genai.list_models()]
        model_name = [m for m in models if "gemini" in m.lower()][0]
        model = genai.GenerativeModel(model_name)
        
        with st.spinner("⏳ Gerando conteúdo com Google Gemini..."):
            prompt = f"""Crie um ebook profissional sobre '{tema_desc}' para '{audience}'.

Estrutura EXATA (importante):
- Título atrativo
- Introdução (2-3 parágrafos)
- {num_chapters} capítulos, cada um com:
  * Título do capítulo
  * 2-3 seções temáticas
  * Conteúdo detalhado (300+ palavras total por capítulo)
- Conclusão
- Dicas práticas

Estilo: {style}
Leve, profissional e prático."""

            response = model.generate_content(prompt)
            content = response.text
            
            st.success("✅ Ebook gerado com sucesso!")
            
            # Exibir conteúdo
            st.markdown("---")
            st.write(content)
            
            # Gerar PDF
with st.spinner("📄 Gerando PDF..."):
    pdf = FPDF()
    pdf.add_font('DejaVu', '', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
    pdf.set_font("DejaVu", "", 11)
    pdf.add_page()
    pdf.set_font("DejaVu", "B", 20)
    pdf.cell(0, 10, tema_nome.replace("🤰", "").replace("👶", "").replace("💪", "").replace("📚", "").replace("💰", "").replace("🧘", "").replace("💼", "").strip(), ln=True, align="C")
    
    pdf.set_font("DejaVu", "I", 10)
    pdf.cell(0, 10, "Por Luciana Britto | L&B Marketing", ln=True, align="C")
    pdf.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y')}", ln=True, align="C")
    
    pdf.ln(10)
    pdf.set_font("DejaVu", "", 11)
    
    # Quebrar texto em linhas
    for line in content.split('\n'):
        if line.strip():
            try:
                pdf.multi_cell(0, 5, line[:500])  # Limita caracteres
            except:
                pass
    
    pdf.ln(5)
    pdf.set_font("DejaVu", "I", 8)
    pdf.cell(0, 5, "© 2026 Luciana Britto | L&B Marketing - Estrategias de Valor", ln=True)
    
    # Salvar em bytes
    pdf_bytes = pdf.output()
                
                # Quebrar texto em linhas
                for line in content.split('\n'):
                    if line.strip():
                        pdf.multi_cell(0, 5, line)
                
                pdf.ln(5)
                pdf.set_font("Arial", "I", 8)
                pdf.cell(0, 5, "© 2026 Luciana Britto | L&B Marketing — Estratégias de Valor", ln=True)
                
                # Salvar em bytes
                pdf_bytes = pdf.output()
                
                st.download_button(
                    label="📥 Download PDF",
                    data=pdf_bytes,
                    file_name=f"{tema_nome.replace('/', '')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            
            # Seção de página de venda
            st.markdown("---")
            st.subheader("🛍️ Próximo Passo: Página de Venda")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.info("""
                **Para vender este ebook:**
                1. Crie página de venda em Canva
                2. Integre Stripe para pagamentos
                3. Configure email automático
                4. Anuncie no Google Ads/Facebook
                """)
            
            with col2:
                st.warning("""
                **Ou use seu skill ebook-architect para:**
                - Validar nicho automaticamente
                - Gerar ebook completo
                - Criar landing page
                - Integrar com Stripe
                """)
            
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
**Luciana Britto | L&B Marketing — Estratégias de Valor**

💡 Dica: Use este gerador para:
- Criar produtos digitais rápido
- Testar nichos antes de investir
- Gerar conteúdo para email marketing
- Criar lead magnets grátis
""")
