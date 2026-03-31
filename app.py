import streamlit as st
import requests

# Page config
st.set_page_config(page_title="BA Resume Pro - Test", page_icon="📊")

st.title("📊 Business Analyst Job Apply Pro")
st.success("✅ Application is running successfully!")

# Sidebar
with st.sidebar:
    st.header("Configuration")
    
    # API Key check
    if "OPENROUTER_API_KEY" in st.secrets:
        api_key = st.secrets["OPENROUTER_API_KEY"]
        st.success("✅ API Key loaded from secrets")
    else:
        api_key = st.text_input("OpenRouter API Key", type="password")
        if api_key:
            st.success("✅ API Key entered")

st.markdown("---")

# Simple test
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Resume")
    uploaded_file = st.file_uploader("Choose file", type=['pdf', 'txt'])
    
with col2:
    st.subheader("2. Job Description")
    jd_text = st.text_area("Paste JD", height=150)

if st.button("🔍 Test Validation"):
    if uploaded_file and jd_text:
        st.success("✅ Both resume and JD provided!")
        st.info(f"📄 File: {uploaded_file.name}")
        st.info(f"📝 JD Length: {len(jd_text)} characters")
        
        # Test API call
        if api_key:
            with st.spinner("Testing API connection..."):
                try:
                    # Simple test request
                    url = "https://openrouter.ai/api/v1/chat/completions"
                    headers = {
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": "application/json"
                    }
                    payload = {
                        "model": "qwen/qwen-2.5-72b-instruct",
                        "messages": [
                            {"role": "user", "content": "Say hello"}
                        ],
                        "max_tokens": 10
                    }
                    response = requests.post(url, headers=headers, json=payload, timeout=30)
                    
                    if response.status_code == 200:
                        st.success("✅ API connection successful!")
                        st.json(response.json())
                    else:
                        st.error(f"❌ API Error: {response.status_code}")
                        st.text(response.text)
                        
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
        else:
            st.warning("⚠️ Please provide API key")
    else:
        st.warning("⚠️ Please upload resume and paste JD")

st.markdown("---")
st.info("💡 This is a minimal test version. If you see this, the app is working!")
