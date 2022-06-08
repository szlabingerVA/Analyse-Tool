import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
from pathlib import Path

st.set_page_config(page_title="Analyse Dashboard", page_icon=":bar_chart:", layout="centered")
#st.title(":bar_chart: Effizienz-Analyse")
#st.markdown("##")

px.colors.qualitative.VoestGrey=["#E3E3E3","#C4C4C4","#A5A5A5"]
px.colors.qualitative.VoestBlue=["#91C8DC","#50AACD","#0082B4"]

col01, col02, col03 = st.columns((1,6,1))
with col02:
    st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')

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
    options=("Effizienz-Analyse","ABC-Analyse","Overview"),
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

col1, col2 = st.columns((3,1))

if selected_hor == "Overview":
    
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">Overview</p>', unsafe_allow_html=True)

    with col2:
        #st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    with st.expander("Upload"):
        files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])
    #file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name)

    if files is not None:

        with st.sidebar:
                st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
                selected_side=option_menu(
                menu_title="Materialgruppe",
                options=("Allgemein","Fertigungsteile","Zukaufteile","Normteile"),
                icons=("caret-right-fill", "caret-right-fill", "caret-right-fill", "caret-right-fill"),
                default_index=0,
                menu_icon="cast",
                styles={
                "container": {"padding": "0!important", "background-color": "#262730"},
                "icon": {"color": "#FAFAFA"}, 
                "nav-link": {"font_family": "voestalpine", "text-align": "left", "--hover-color": "#A5A5A5"},
                "nav-link-selected": {"background-color": "#0082B4"},
                }
                )

                if selected_side == "Fertigungsteile":
                    sheet="Komponenten_FERT"
                if selected_side == "Zukaufteile":
                    sheet="Komponenten_HIBE"
                if selected_side == "Normteile":
                    sheet="Komponenten_NORM"
                if selected_side == "Allgemein":
                    sheet="Komponenten"

        dfs={}
        for file in files:
            dfs[f"{file.name}"]=pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name=sheet,
                            na_filter=False)

        fill=[]
        for file in files:
            for i in range(len(dfs[file.name].index)):
                fill.append(i/(len(dfs[file.name].index)))
            dfs[file.name]["Prozent"]=fill
            fill.clear()
        
        fig = go.Figure()

        for i in dfs:
             fig = fig.add_trace(go.Scatter(x = dfs[i]["Prozent"],
                                   y = dfs[i]["Effizienz"], 
                                   name = i,
                                   mode='lines'))
        
        fig.update_layout(
            plot_bgcolor="rgba(256,256,256,0)",
            title_font_family="voestalpine",
            font_family="voestalpine",
            xaxis=(dict(showgrid=False,
                        showticklabels=False)),
            yaxis=(dict(showgrid=False,
                        tickfont_size=20,
                        titlefont_size=30)),
            yaxis_tickformat = ',.0%'
            )
        
        st.plotly_chart(fig, use_container_width=True, height=800)

if selected_hor == "ABC-Analyse":
    
    with col1:
        st.markdown("""
            <style>
            .big-font {
            font-size:50px !important;
            color: #0082B4
            }
            </style>
            """, unsafe_allow_html=True)
        st.markdown('<p class="big-font">ABC-Analyse</p>', unsafe_allow_html=True)

    with col2:
        #st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    with st.expander("Upload"):
        files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])
    file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name)

    if file is not None:
        dfr = pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name="ABC",
                        na_filter=False,
                        usecols="A:D")

        dfr['Wertanteil'] = dfr['Wertanteil'].map('{:,.1f}%'.format)
        dfr['Summe'] = dfr['Summe'].map('{:,.1f}%'.format)
        
        df = pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name="ABC",
                        na_filter=False,
                        usecols="E:F",
                        header=None)

        fill=[]
        for i in range(len(dfr.index)+1):
            fill.append("-")
        
        df[6]=fill

        df.at[0,6]="Klasse A"
        df.at[1,6]="Klasse B"
        df.at[2,6]="Klasse C"

        fig_bar1 = px.bar(df, y=4, color=6, text=4, 
            barmode = 'stack', orientation="v",
            labels={"count":"Mengenanteil", "4":""},
            color_discrete_sequence=px.colors.qualitative.VoestBlue)
        
        fig_bar2 = px.bar(df, y=5, color=6, text=5,
            barmode = 'stack', orientation="v",
            labels={"count":"Wertanteil", "5":""},
            color_discrete_sequence=px.colors.qualitative.VoestBlue)

        fig_bar1.update_layout(
            plot_bgcolor="rgba(256,256,256,0)",
            title_font_family="voestalpine",
            font_family="voestalpine",
            showlegend=False,
            xaxis=(dict(showgrid=False,
                        showticklabels=False,
                        titlefont_size=30)),
            yaxis=(dict(showgrid=False,
                        tickfont_size=15)),
            hoverlabel=dict(
                bgcolor="#0082B4",
                font_size=16,
                font_family="voestalpine"
                )
        )

        fig_bar2.update_layout(
            plot_bgcolor="rgba(256,256,256,0)",
            title_font_family="voestalpine",
            font_family="voestalpine",
            showlegend=False,
            xaxis=(dict(showgrid=False,
                        showticklabels=False,
                        titlefont_size=30)),
            yaxis=(dict(showgrid=False,
                        tickfont_size=15)),
            hoverlabel=dict(
                bgcolor="#0082B4",
                font_size=16,
                font_family="voestalpine"
                )
        )

        fig_bar1.update_traces(marker=dict(line=dict(width=0)),
            hovertemplate='%{y:.2f}%'+"<br>Mengenanteil",
            texttemplate='%{text:.2f}%',
            textposition="none",
            textfont=dict(
                family="voestalpine",
                size=25,))
        fig_bar2.update_traces(marker=dict(line=dict(width=0)),
            hovertemplate='%{y:.2f}%'+"<br>Wertanteil",
            texttemplate='%{text:.2f}%',
            textposition="none",
            textfont=dict(
                family="voestalpine",
                size=25,))
        
        col11, col12, col13 = st.columns((4,2,4))

        df[4].iloc[:3] = df[4].iloc[:3].map('{:,.1f}%'.format)
        df[5].iloc[:3] = df[5].iloc[:3].map('{:,.1f}%'.format)

        with col11:
            st.plotly_chart(fig_bar1, use_container_width=True)

        with col12:
            st.write("")
            st.write("")
            st.write("")
            st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Klasse C</p>', unsafe_allow_html=True)
            st.write(df.at[2,4]+" :arrow_right: "+df.at[2,5])
            st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Klasse B</p>', unsafe_allow_html=True)
            st.write(df.at[1,4]+" :arrow_right: "+df.at[1,5])
            st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Klasse A</p>', unsafe_allow_html=True)
            st.write(df.at[0,4]+" :arrow_right: "+df.at[0,5])


        with col13:
            st.plotly_chart(fig_bar2, use_container_width=True)
        
        col21, col22, col23 = st.columns((1,6,1))

        with col22:
            st.dataframe(dfr, height=400, width=600)

if selected_hor == "Effizienz-Analyse":

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

    with st.expander("Upload"):
        files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])
    file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name)

    if file is not None:

        with st.sidebar:
            st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
            selected_side=option_menu(
            menu_title="Materialgruppe",
            options=("Allgemein","Fertigungsteile","Zukaufteile","Normteile"),
            icons=("caret-right-fill", "caret-right-fill", "caret-right-fill", "caret-right-fill"),
            default_index=0,
            menu_icon="cast",
            styles={
            "container": {"padding": "0!important", "background-color": "#262730"},
            "icon": {"color": "#FAFAFA"}, 
            "nav-link": {"font_family": "voestalpine", "text-align": "left", "--hover-color": "#A5A5A5"},
            "nav-link-selected": {"background-color": "#0082B4"},
            }
            )

            if selected_side == "Fertigungsteile":
                sheet="Komponenten_FERT"
            if selected_side == "Zukaufteile":
                sheet="Komponenten_HIBE"
            if selected_side == "Normteile":
                sheet="Komponenten_NORM"
            if selected_side == "Allgemein":
                sheet="Komponenten"

        dfd=pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name="Eingabe",
                        na_filter=False,
                        usecols="A")

        df = pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name=sheet,
                        na_filter=False)

        var=len(dfd.index)
        ein=len(df.index)
        unicount=0
        exklcount=0

        for i in range(len(df.index)):
            if df.at[i,"Effizienz"] == 1:
                unicount=unicount+1

        for i in range(len(df.index)):
            if df.at[i,"Abfrage"] == 1:
                exklcount=exklcount+1

        uni=unicount
        exkl=exklcount

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

        
        st.plotly_chart(fig_area, use_container_width=True)

        col11, col12, col13 = st.columns((2,3,6))

        st.markdown("""
            <style>
            .font {
           font:30px voestalpine;
            }
            </style>
            """, unsafe_allow_html=True)

        unistr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{uni}</p>"""
        exklstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{exkl}</p>"""
        varstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{var}</p>"""
        einstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{ein}</p>"""

        with col11:
            st.markdown(unistr, unsafe_allow_html=True)
            st.markdown(exklstr, unsafe_allow_html=True)
            st.markdown(varstr, unsafe_allow_html=True)
            st.markdown(einstr, unsafe_allow_html=True)
        
        with col12:
            st.markdown('<p class="font">Universalkomp.</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Exklusivkomp.</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Varianten</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Einzelkomp.</p>', unsafe_allow_html=True)

        df = df.iloc[1: , :]
        
        with col13:
            st.dataframe(df.loc[:, ["Komponentennummer", "Effizienz", "Materialart", "Objektsparte"]], height=210)
