import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_option_menu import option_menu
from pathlib import Path
from PIL import Image

st.set_page_config(page_title="Analyse Dashboard", page_icon=":bar_chart:", layout="centered")
#st.title(":bar_chart: Effizienz-Analyse")
#st.markdown("##")

st.markdown(
        """
        <style>
@font-face {
  font-family: 'voestalpine';
  font-style: normal;
  font-weight: 400;
}
    html, body, [class*="css"]  {
    font-family: 'voestalpine';
    font-size: 20px;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

selected_hor=option_menu(
    menu_title=None,
    options=("Analyse","Upload","Overview"),
    icons=("bar-chart-fill", "cloud-arrow-up", "bezier"),
    default_index=0,
    menu_icon="cast",
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#262730"},
        "icon": {"color": "#FAFAFA"}, 
        "nav-link": {"font_family": "voestalpine", "text-align": "center", "--hover-color": "#A5A5A5"},
        "nav-link-selected": {"background-color": "#0082B4"},
    }
)
with st.sidebar:
    selected_hor=option_menu(
        menu_title="Filter",
        options=("Materialgruppe","Effizienz","Objektsparte"),
        icons=("caret-right-fill", "caret-right-fill", "caret-right-fill"),
        default_index=0,
        menu_icon="cast",
        styles={
            "container": {"padding": "0!important", "background-color": "#262730"},
            "icon": {"color": "#FAFAFA"}, 
            "nav-link": {"font_family": "voestalpine", "text-align": "left", "--hover-color": "#A5A5A5"},
            "nav-link-selected": {"background-color": "#0082B4"},
        }
    )

col1, col2 = st.columns((3,1))

with col1:
    st.markdown("""
        <style>
        .big-font {
         font-size:50px !important;
         color: #0082B4
        }
        </style>
        """, unsafe_allow_html=True)
    st.markdown('<p class="big-font">Effizienz-Analyse</p>', unsafe_allow_html=True)

with col2:
    #st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
    st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

my_expander = st.expander("UPLOADER", expanded=True)
with my_expander:
    files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])

file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name)

if file is not None:
    df = pd.read_excel(file,
                    engine="openpyxl",
                    sheet_name="Komponenten")

    df["Effizienz"] = df["Effizienz"]*100
    df['Effizienz'] = df['Effizienz'].map('{:,.1f}%'.format)

    df["Mittelwert"] = df["Mittelwert"]*100
    df['Mittelwert'] = df['Mittelwert'].map('{:,.1f}%'.format)

    df.iloc[0] = ['-','-','-','-','-','-','-','-','-','-','-','-',"0","0",'-','-']

    fig_area = px.bar(
        data_frame=df, y="Effizienz", x="Komponentennummer",
        labels=dict(Komponentennummer="", Effizienz="Bauteileffizienz"), 
        height=500,
        range_y=(0,100),
        base=None,
        hover_name="Objektsparte",
        #color="Abfrage",
        #color_continuous_scale=px.colors.sequential.Oryel
        #color_discrete_sequence=px.colors.sequential.Plasma_r
        color_discrete_sequence =['#0082B4']*len(df)
    )

    fig_area.update_layout(
        plot_bgcolor="rgba(256,256,256,0)",
        title_font_family="voestalpine",
        font_family="voestalpine",
        xaxis=(dict(showgrid=False,
                    showticklabels=False)),
        yaxis=(dict(showgrid=False,
                    tickfont_size=20,
                    titlefont_size=30)),
    )
    fig_area['layout']['yaxis'].update(autorange = True)

    fig_area.update_traces(marker=dict(line=dict(width=0)))

    col11, col12, col13 = st.columns([1,6,1])

    #with col12:
    st.plotly_chart(fig_area, use_container_width=True)

    df = df.iloc[1: , :]

    with col12:
        st.write(df.loc[:, ["Komponentennummer", "Effizienz", "Materialart", "Objektsparte"]], height=400)

