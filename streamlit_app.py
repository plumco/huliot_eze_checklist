import streamlit as st
import streamlit.components.v1 as components
import os

# Set page configurations to clean full-screen widescreen mode
st.set_page_config(
    page_title="Huliot Eze Checklist Workstation",
    page_icon="☑",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Hide Streamlit's default headers and footers to look like a premium standalone app
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    div.block-container {padding: 0px;}
    iframe {border: none; width: 100vw; height: 100vh;}
    body {overflow: hidden;}
    </style>
""", unsafe_allow_html=True)

def main():
    html_filename = "index.html"
    
    # Check if index.html exists in your root folder
    if os.path.exists(html_filename):
        with open(html_filename, "r", encoding="utf-8") as f:
            html_content = f.read()
        
        # Render the full HTML workspace with complete browser permissions enabled
        components.html(
            html_content,
            height=1000,
            scrolling=True
        )
    else:
        st.error(f"Error: '{html_filename}' not found in the repository root directory. Please check your GitHub files.")

if __name__ == "__main__":
    main()
