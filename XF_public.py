import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from PIL import Image

st.set_page_config(page_title="Analyse Dashboard", page_icon=":bar_chart:", layout="wide")
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

#names = files.split("/")
#names = [i.split("/")[0] for i in files]
#names = [val[-1] for val in str(files).split("/")]

#voest=Image.open("U:/WORK/Logos und Vorlagen/voestalpine_LOGO.jpg")
col1, col2, col3, col4 = st.columns(4)

with col1:
    #st.title(":bar_chart: Effizienz-Analyse")
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
    st.write("")

with col3:
    st.write("")

with col4:
    st.write("")

files = st.file_uploader("Upload", accept_multiple_files=True, type=["xlsm"])

file=st.selectbox("Analyse-Files", files)

df = pd.read_excel(file,
                   engine="openpyxl",
                  sheet_name="Komponenten")

df["Effizienz"] = df["Effizienz"]*100
df['Effizienz'] = df['Effizienz'].map('{:,.1f}%'.format)

df.iloc[0] = ['-','-','-','-','-','-','-','-','-','-','-','-',"0","0",'-','-']

col11, col21 = st.columns(2)

fig_area = px.bar(
    data_frame=df, y="Effizienz", x="Komponentennummer",
    labels=dict(Komponentennummer="", Effizienz="Bauteileffizienz"), 
    height=500,
    width=1000,
    range_y=(0,100),
    base=None,
    #color="Abfrage",
    #color_continuous_scale=px.colors.sequential.Oryel
    #color_discrete_sequence=px.colors.sequential.Plasma_r
    color_discrete_sequence =['#0082B4']*len(df)
)

fig_area.update_layout(
    plot_bgcolor="rgba(256,256,256,1)",
    title_font_family="voestalpine",
    font_family="voestalpine",
    xaxis=(dict(showgrid=False,
                showticklabels=False)),
    yaxis=(dict(showgrid=False,
                tickfont_size=20,
                titlefont_size=30)),
)
fig_area['layout']['yaxis'].update(autorange = True)

with col21:
    st.plotly_chart(fig_area)

df = df.iloc[1: , :]

with col11:
    st.dataframe(df.loc[:, ["Komponentennummer", "Effizienz", "Materialart", "Objektsparte"]], height=400)
