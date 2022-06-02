import streamlit as st
import pandas as pd
import plotly.express as px
import os
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

    html, h1, [class*="css"]  {
    font-family: 'voestalpine';
    font-size: 30px;
    }

     html, body, [class*="css"]  {
    font-family: 'voestalpine';
    font-size: 20px;
    }

    </style>

    """,
        unsafe_allow_html=True,
    )

INPUT = Path('V:/02_DLDS_Operational_Workplace/04_Engineering/Variantenanalyse/')
files = list(INPUT.rglob("*Analyse.xls*"))

#names = files.split("/")
#names = [i.split("/")[0] for i in files]
#names = [val[-1] for val in str(files).split("/")]

#voest=Image.open("U:/WORK/Logos und Vorlagen/voestalpine_LOGO.jpg")

#st.title(":bar_chart: Effizienz-Analyse")
st.markdown("""
        <style>
        .big-font {
         font-size:50px !important;
        }
        </style>
        """, unsafe_allow_html=True)
st.markdown('<p class="big-font">Effizienz-Analyse</p>', unsafe_allow_html=True)

names=[]
for i in files:
    names.append(i.name)

parts=[]
for path in files:
    part = pd.read_excel(path, 
    engine="openpyxl",
    sheet_name="Komponenten")
    parts.append(part)

st.write(str(files))
        
option=st.selectbox("Analyse-Files", names)
result=[]
result=list(INPUT.rglob(option))

resultpath=[]
for j in result:
    resultpath.append(j.name)

#st.text_area(str(resultpath))

fpath = os.path.dirname(j)

df = pd.read_excel(
    io=fpath,
    engine="openpyxl",
    sheet_name="Komponenten"
    )

df["Effizienz"] = df["Effizienz"]*100
df['Effizienz'] = df['Effizienz'].map('{:,.1f}%'.format)

col11, col21 = st.columns(2)

with col11:
    st.dataframe(df.loc[:, ["Komponentennummer", "Effizienz", "Materialart", "Objektsparte"]], height=400)

fig_area = px.bar(
    data_frame=df, y="Effizienz", x="Komponentennummer",
    labels=dict(Komponentennummer="", Effizienz="Bauteileffizienz"), 
    height=500,
    width=1000,
    range_y=(0,100),
    base=None,
    color="Abfrage",
    #color_continuous_scale=px.colors.sequential.Oryel
    color_discrete_sequence=px.colors.sequential.Plasma_r
    #color_discrete_sequence =['blue']*len(df)
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
