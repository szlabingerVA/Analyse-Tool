import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from streamlit_option_menu import option_menu
from plotly.subplots import make_subplots
from streamlit_echarts import st_echarts
from st_aggrid import AgGrid, GridOptionsBuilder

#Webapp Konfiguration
st.set_page_config(page_title="Analyse Dashboard", page_icon=":bar_chart:", layout="centered")

#Farbpaletten definieren
px.colors.qualitative.VoestGrey=["#E3E3E3","#C4C4C4","#A5A5A5"]
px.colors.qualitative.VoestBlue=["#91C8DC","#50AACD","#0082B4"]

#WebApp Kopf erstellen
col01, col02, col03 = st.columns((1,6,1))
with col02:
    st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')

#Globale Schriftart definieren
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

#Horizontales Optionsmenü erstellen
selected_hor=option_menu(
    menu_title=None,
    options=("Overview","Effizienz-Analyse","ABC-Analyse","Portfolio"),
    icons=("bezier", "bar-chart-fill", "calculator-fill", "grid-3x3-gap-fill"),
    default_index=1,
    menu_icon="cast",
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#262730"},
        "icon": {"color": "#FAFAFA"}, 
        "nav-link": {"font_family": "voestalpine", "text-align": "center", "--hover-color": "#A5A5A5"},
        "nav-link-selected": {"background-color": "#0082B4"},
    }
)

#Spalten definieren
col1, col2 = st.columns((3,1))

#Option Overview
if selected_hor == "Overview":
    
    #Überschrift erstellen
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

    #Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    #Datei-Uploader einfügen
    with st.expander("Upload"):
        filesInput = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])

    #Dashboard bei Datei-Upload
    if filesInput is not None:
        
        try:
            #Sidebar definieren
            with st.sidebar:
                    #VoestAlpine Logo einfügen 
                    st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
                    #Sidebar Optionsmenü erstellen
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
                    
                    #Auswahl zuordnen
                    if selected_side == "Fertigungsteile":
                        sheet="Komponenten_FERT"
                    if selected_side == "Zukaufteile":
                        sheet="Komponenten_HIBE"
                    if selected_side == "Normteile":
                        sheet="Komponenten_NORM"
                    if selected_side == "Allgemein":
                        sheet="Komponenten"
                    
                    #Overview Auswahl
                    files=st.multiselect("Baugruppen-Auswahl", filesInput, default=filesInput, format_func=lambda x: x.name.split("_")[0]+x.name.split("_")[1])

            #Ausgewählte Dataframes kombinieren
            dfs={}
            for file in files:
                dfs[f"{file.name}"]=pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name=sheet,
                                na_filter=True)

            #Spalte hinzufügen für Skalierung der X-Achse
            fill=[]
            for file in files:
                for i in range(len(dfs[file.name].index)):
                    fill.append(i/(len(dfs[file.name].index)))
                dfs[file.name]["Prozent"]=fill
                fill.clear()
            
            #Universal-Anteil-Liste
            unipL=[]
            #Exklusiv-Anteil-Liste
            exklpL=[]

            #Parameter ermitteln
            for file in files:
                #Anzahl der Einzelkomponenten
                ein=len(dfs[file.name].index)
                unicount=0
                exklcount=0

                #Universalkomponenten zählen
                for i in range(ein):
                    if dfs[file.name].at[i,"Effizienz"] == 1:
                        unicount=unicount+1

                #Universalanteil ermitteln und der Liste zuweisen
                unip=unicount/ein
                unipL.append(unip)
                
                #Exklusivkomponenten zählen
                for i in range(ein):
                    if dfs[file.name].at[i,"Abfrage"] == 1:
                        exklcount=exklcount+1

                #Exklusivanteil ermitteln und der Liste zuweisen
                exklp=exklcount/ein
                exklpL.append(exklp)
            
            #Mittelwerte ermitteln
            SumUni=sum(unipL)
            avUni=(SumUni/len(unipL))*100        
            SumExkl=sum(exklpL)
            avExkl=(SumExkl/len(exklpL))*100

            #Fertigungsteile
            dfsPF={}
            for file in files:
                dfsPF[f"{file.name}"]=pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="Komponenten_FERT",
                                na_filter=False)
            
            PF=[]
            for file in files:
                PF.append(len(dfsPF[file.name].index))
            
            SumFERT=sum(PF)
            
            #Zukaufteile
            dfsPH={}
            for file in files:
                dfsPH[f"{file.name}"]=pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="Komponenten_HIBE",
                                na_filter=False)

            PH=[]
            for file in files:
                PH.append(len(dfsPH[file.name].index))

            SumHIBE=sum(PH)

            #Normteile
            dfsPN={}
            for file in files:
                dfsPN[f"{file.name}"]=pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name="Komponenten_NORM",
                                na_filter=False)              

            PN=[]
            for file in files:
                PN.append(len(dfsPN[file.name].index))

            SumNORM=sum(PN)

            #General
            dfsP={}
            for file in files:
                dfsP[f"{file.name}"]=pd.read_excel(file,
                                engine="openpyxl",
                                sheet_name=sheet,
                                na_filter=False)

            #Gesamtsumme
            SumGEN=SumFERT+SumHIBE+SumNORM

            #Anteilermittlung
            AnFERT=SumFERT/SumGEN
            AnHIBE=SumHIBE/SumGEN
            AnNORM=SumNORM/SumGEN
            
            #Mittelwerte ermitteln
            MW=[]
            for i in dfsP:
                MW.append(sum(dfsP[i]["Effizienz"])/len(dfsP[i].index))
            
            #Mittelwerte benennen
            MWN=[]
            for file in files:
                MWN.append(file.name.split("_")[1])

            #Komponenten ohne Materialart entfernen
            for file in files:
                dfs[file.name].dropna(subset=["Materialart"], inplace=True)

            #Dashgrid definieren
            fig = make_subplots(
                rows=3, cols=2,
                specs=[[{},{"type":"domain"}],
                    [{"colspan": 2}, None],
                    [{"type":"indicator"},{"type":"indicator"}]]
            )

            #Balkendiagramm (Avg.Effizienz) einfügen
            fig.add_trace(go.Bar(y=MW, x=MWN), row=1, col=1)
            fig.update_traces(marker_color="#0082B4", 
                            showlegend=False,
                            hovertemplate = "%{label}: <br>%{value}",
                            row=1, col=1)
            fig.update_xaxes(categoryorder='total descending',
                            showticklabels=False, 
                            showgrid=False,
                            row=1, col=1)
            fig.update_yaxes(showgrid=False,
                            title="Avg.Effizienz",
                            tickfont_size=20,
                            titlefont_size=30,
                            titlefont_family="voestalpine",
                            tickformat = ',.0%',
                            row=1, col=1)

            #Tortendiagramm (Materialartanteil) einfügen
            labels = ['FERT','HIBE','NORM']
            values = [AnFERT, AnHIBE, AnNORM]
                    
            fig.add_trace(go.Pie(labels=labels, values=values, hole=.3),row=1,col=2)
            fig.update_traces(marker=dict(colors=px.colors.qualitative.VoestBlue), 
                            showlegend=False,
                            title=" <br>MatGr.Aufteilung",
                            titlefont_size=20,
                            titlefont_family="voestalpine",
                            titleposition="bottom center",
                            textfont_size=18,
                            textposition='inside',
                            insidetextorientation='radial',
                            textfont_family="voestalpine",
                            hoverinfo='label+percent',
                            textinfo='percent+label',
                            row=1, col=2)

            #Liniendiagramm (Bauteileffizienz) einfügen
            for i in dfs:
                fig.add_trace(go.Scatter(x = dfs[i]["Prozent"],
                                    y = dfs[i]["Effizienz"], 
                                    name = i.split("_")[1],
                                    mode='lines'),
                                    row=2,col=1)

            fig.update_xaxes(showgrid=False,
                            showticklabels=False,
                            row=2, col=1)
            fig.update_yaxes(showgrid=False,
                            title="Bauteileffizienz",
                            tickfont_size=20,
                            titlefont_size=30,
                            titlefont_family="voestalpine",
                            tickformat = ',.0%',
                            row=2, col=1)        
            fig.update_traces(showlegend=False,
                            hoverinfo="name+y",
                            hoverlabel_namelength=-1,
                            row=2, col=1)

            #Zeigerdiagramm (Universalanteil) einfügen
            fig.add_trace(go.Indicator(domain = {'x': [0, 1], 'y': [0, 1]},
                    value = avUni,
                    mode = "gauge+number",
                    title = {'text': "Universalanteil", 'font': {'size': 40, 'family': 'voestalpine'}},
                    #delta = {'reference': 50},
                    number = {'suffix': "%"},
                    gauge = {
                        "bar":{"color": "#0eff00"},
                        'axis': {'range': [None, 100]}
                    }),
                    row=3, col=1)

            #Zeigerdiagramm (Exklusivanteil) einfügen
            fig.add_trace(go.Indicator(domain = {'x': [0, 1], 'y': [0, 1]},
                    value = avExkl,
                    mode = "gauge+number",
                    title = {'text': "Exklusivanteil", 'font': {'size': 40, 'family': 'voestalpine'}},
                    #delta = {'reference': 50},
                    number = {'suffix': "%"},
                    gauge = {
                        "bar":{"color": "red"},
                        'axis': {'range': [None, 100]}
                    }),
                    row=3, col=2)

            #Dashlayout anpassen
            fig.update_layout(plot_bgcolor="rgba(256,256,256,0)",
                            legend_font=dict(family="voestalpine", size=18),
                            legend=dict(
                                orientation="h",
                                yanchor="top",
                                y=0.74,
                                xanchor="left",
                                x=0.01,
                                bgcolor="rgba(256,256,256,0)"),
                            height=1000)
            
            #Dash anzeigen
            st.plotly_chart(fig, use_container_width=True)

            st.stop()

        #Bei fehlenden Input:
        except ZeroDivisionError:
            st.write("Bitte uploaden Sie die Analyse-Dateien!")

#Option ABC-Analyse
if selected_hor == "ABC-Analyse":
    
    try:
        #Überschrift erstellen
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

        #Railway Systems Logo einfügen
        with col2:
            st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

        #Datei-Uploader einfügen
        with st.expander("Upload"):
            files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])
        
        files.sort(key=lambda x: x.name.split("_")[1])

        #Selectbox einfügen
        file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name.split("_")[0]+x.name.split("_")[1])

        #Bei Datei-Upload
        if file is not None:
            dfr = pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name="ABC",
                            na_filter=False,
                            usecols="A:D")

            #Spalten formatieren
            dfr['Wertanteil'] = dfr['Wertanteil'].map('{:,.1f}%'.format)
            dfr['Summe'] = dfr['Summe'].map('{:,.1f}%'.format)
            
            #ABC-Dataframe öffnen
            df = pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name="ABC",
                            na_filter=False,
                            usecols="E:F",
                            header=None)

            #Spalte hinzufügen
            fill=[]
            for i in range(len(dfr.index)+1):
                fill.append("-")
            
            df[6]=fill

            #Ersten drei Zeilen überschreiben
            df.at[0,6]="Klasse A"
            df.at[1,6]="Klasse B"
            df.at[2,6]="Klasse C"

            #Balkendiagramm (Mengenanteil) erstellen
            fig_bar1 = px.bar(df, y=4, color=6, text=4, 
                barmode = 'stack', orientation="v",
                labels={"count":"Mengenanteil", "4":""},
                color_discrete_sequence=px.colors.qualitative.VoestBlue)
            
            #Balkendiagramm (Wertanteil) erstellen
            fig_bar2 = px.bar(df, y=5, color=6, text=5,
                barmode = 'stack', orientation="v",
                labels={"count":"Wertanteil", "5":""},
                color_discrete_sequence=px.colors.qualitative.VoestBlue)

            #Dashlayout 1 anpassen
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

            #Dashlayout 2 anpassen
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
            
            #Dashtraces 1 anpassen
            fig_bar1.update_traces(marker=dict(line=dict(width=0)),
                hovertemplate='%{y:.2f}%'+"<br>Mengenanteil",
                texttemplate='%{text:.2f}%',
                textposition="none",
                textfont=dict(
                    family="voestalpine",
                    size=25,))
            
            #Dashtraces 2 anpassen
            fig_bar2.update_traces(marker=dict(line=dict(width=0)),
                hovertemplate='%{y:.2f}%'+"<br>Wertanteil",
                texttemplate='%{text:.2f}%',
                textposition="none",
                textfont=dict(
                    family="voestalpine",
                    size=25,))
            
            #Spalten definieren
            col11, col12, col13 = st.columns((4,2,4))

            #Zellen formatieren
            df[4].iloc[:3] = df[4].iloc[:3].map('{:,.1f}%'.format)
            df[5].iloc[:3] = df[5].iloc[:3].map('{:,.1f}%'.format)

            #Balkendiagramm (Mengenanteil) anzeigen
            with col11:
                st.plotly_chart(fig_bar1, use_container_width=True)

            #Werte anzeigen
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

            #Balkendiagramm (Wertanteil) anzeigen
            with col13:
                st.plotly_chart(fig_bar2, use_container_width=True)
            
            #Spalten definieren
            col21, col22, col23 = st.columns((1,6,1))

            #ABC-Dataframe anzeigen
            with col22:
                st.dataframe(dfr, height=400, width=600)
            
            st.stop()

    except ValueError:
        st.write("Die Input-Datei hat keine passende ABC-Aufschlüsselung!")

#Option Effizienz-Analyse
if selected_hor == "Effizienz-Analyse":

    #Überschrift einfügen
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

    #Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    #Datei-Uploader einfügen
    with st.expander("Upload"):
        files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])

    files.sort(key=lambda x: x.name.split("_")[1])

    #Selectbox einfügen
    file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name.split("_")[0]+x.name.split("_")[1])

    #Bei Datei-Uplaod
    if file is not None:
        
        #Sidebar erstellen
        with st.sidebar:
            #VoestAlpine Logo einfügen
            st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
            #Sidebar Optionsmenü erstellen
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

            #Auswahl zuordnen
            if selected_side == "Fertigungsteile":
                sheet="Komponenten_FERT"
            if selected_side == "Zukaufteile":
                sheet="Komponenten_HIBE"
            if selected_side == "Normteile":
                sheet="Komponenten_NORM"
            if selected_side == "Allgemein":
                sheet="Komponenten"

        #Eingabe-Dataframe öffnen
        dfd=pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name="Eingabe",
                        na_filter=False,
                        usecols="A")

        #Ausgewählten Dataframe öffnen
        df = pd.read_excel(file,
                        engine="openpyxl",
                        sheet_name=sheet,
                        na_filter=True)

        #Parameter ermitteln und initialisieren
        var=len(dfd.index)
        ein=len(df.index)
        unicount=0
        exklcount=0
        
        #Universalkomponenten zählen
        for i in range(len(df.index)):
            if df.at[i,"Effizienz"] == 1:
                unicount=unicount+1

        #Exklusivkomponenten zählen
        for i in range(len(df.index)):
            if df.at[i,"Abfrage"] == 1:
                exklcount=exklcount+1

        uni=unicount
        exkl=exklcount

        #Anteile ermitteln
        unip=(unicount/ein)*100
        exklp=(exklcount/ein)*100

        #Spalte Effizienz formatieren
        df["Effizienz"] = df["Effizienz"]*100
        df['Effizienz'] = df['Effizienz'].map('{:,.1f}%'.format)

        #Spalte Mittelwert formatieren
        df["Mittelwert"] = df["Mittelwert"]*100
        df['Mittelwert'] = df['Mittelwert'].map('{:,.1f}%'.format)

        #Zeile einfügen
        df.iloc[0] = ['-','-','-','-','-','-','-','-','-','-','-','-',"0","0",'-','-']

        #Balkendiagramm (Bauteileffizienz) erstellen
        fig_area = px.bar(
            data_frame=df, y="Effizienz", x="Komponentennummer",
            labels=dict(Komponentennummer="Komponentenanzahl", Effizienz="Bauteileffizienz"), 
            height=500,
            range_y=(0,100),
            base=None,
            hover_name="Objektsparte",
            #hover_data=["Abfrage"],
            custom_data=["Objektsparte","Abfrage"],
            #color="Abfrage",
            #color_continuous_scale=px.colors.sequential.Oryel
            #color_discrete_sequence=px.colors.sequential.Plasma_r
            color_discrete_sequence =['#0082B4']*len(df)
        )

        fig_area.update_traces(
            hovertemplate=
                "<b>%{customdata[0]}</b><br>"+
                "Mat.Nr.: <b>%{x}</b><br>"+
                "Bauteileffizienz: <b>%{y}</b><br>"+
                "In <b>%{customdata[1]}</b> Varianten verbaut"   
        )

        #Dashlayout anpassen
        fig_area.update_layout(
            plot_bgcolor="rgba(256,256,256,0)",
            title_font_family="voestalpine",
            font_family="voestalpine",
            xaxis=(dict(showgrid=False,
                        showticklabels=False,
                        tickfont_size=20,
                        titlefont_size=30)),
            yaxis=(dict(showgrid=False,
                        tickfont_size=20,
                        titlefont_size=30)),
        )

        #Y-Achse anpassen
        fig_area['layout']['yaxis'].update(autorange = True)

        #Dashtraces anpassen
        fig_area.update_traces(marker=dict(line=dict(width=0)))
        
        #Komponenten ohne Materialart entfernen
        df.dropna(subset=["Materialart"], inplace=True)

        #Dash anzeigen
        st.plotly_chart(fig_area, use_container_width=True)

        #Spalten definieren
        col11, col12, col13 = st.columns((2,3,6))

        #Schriftart definieren
        st.markdown("""
            <style>
            .font {
        font:30px voestalpine;
            }
            </style>
            """, unsafe_allow_html=True)

        #Textformatierung zuweisen
        unistr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{uni}</p>"""
        exklstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{exkl}</p>"""
        varstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{var}</p>"""
        einstr = f"""<style>p.a {{font:30px voestalpine;color: #0082B4;text-align: right}}</style><p class="a">{ein}</p>"""

        #Werte anzeigen
        with col11:
            st.markdown(unistr, unsafe_allow_html=True)
            st.markdown(exklstr, unsafe_allow_html=True)
            st.markdown(varstr, unsafe_allow_html=True)
            st.markdown(einstr, unsafe_allow_html=True)
        
        #Beschriftung anzeigen
        with col12:
            st.markdown('<p class="font">Universalkomp.</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Exklusivkomp.</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Varianten</p>', unsafe_allow_html=True)
            st.markdown('<p class="font">Einzelkomp.</p>', unsafe_allow_html=True)

        #Dataframe anpassen
        df = df.iloc[1: , :]
        
        gb = GridOptionsBuilder.from_dataframe(df)
        
        grid_options = {
            "defaultColDef": {
                "minWidth": 5,
                "editable": False,
                "filter": True,
                "resizable": True,
                "sortable": True
            },
            "suppressFieldDotNotation": True,
            "columnDefs": [
                {
                "headerName": "Komponentennummer",
                "field": "Komponentennummer",
                "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"15px"},
                "width": 280,
                "type": []
                },
                {
                "headerName": "Effizienz",
                "field": "Effizienz",
                "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"15px"},
                "width": 130,
                "type": []
                },
                {
                "headerName": "Materialart",
                "field": "Materialart",
                "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"15px"},
                "width": 130,
                "type": []
                },
                {
                "headerName": "Objektkurztext",
                "field": "Objektkurztext",
                "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"15px"},
                "width": 370,                
                "type": []
                },
            ]
            }
        
        #Dataframe anzeigen
        with col13:
           (AgGrid(df,gridOptions=grid_options, theme="streamlit", height= 210,fit_columns_on_grid_load=True, editable=False))
            
        #Spalten definieren
        col111, col112, col113 = st.columns((8,5,8))

        #Zeigerdiagramm (Universalanteil) erstellen
        fig_uni = go.Figure(go.Indicator(
            domain = {'x': [0, 1], 'y': [0, 1]},
            value = unip,
            mode = "gauge+number",
            title = {'text': "Universalanteil", 'font': {'size': 40, 'family': 'voestalpine'}},
            #delta = {'reference': 50},
            number = {'suffix': "%"},
            gauge = {
                    "bar":{"color": "#0eff00"},
                    'axis': {'range': [None, 100]}
                    }))

        #Zeigerdiagramm (Exklusivanteil) erstellen
        fig_exkl = go.Figure(go.Indicator(
            domain = {'x': [0, 1], 'y': [0, 1]},
            value = exklp,
            mode = "gauge+number",
            title = {'text': "Exklusivanteil", 'font': {'size': 40, 'family': 'voestalpine'}},
            #delta = {'reference': 50},
            number = {'suffix': "%"},
            gauge = {
                    "bar":{"color": "red"},
                    'axis': {'range': [None, 100]}
                    }))
        
        #Zeigerdiagramm (Universalanteil) anzeigen
        with col111:
            st.plotly_chart(fig_uni, use_container_width=True, height=500)

        #Zeigerdiagramm (Exklusivanteil) anzeigen
        with col113:
            st.plotly_chart(fig_exkl, use_container_width=True)

        #Häufigsten Bauteile ermitteln
        n = 5
        topl=df['Objektsparte'].value_counts().index.tolist()[:n]
        topl.insert(0, "test")
        data= {"Komponenten":topl}
        dftop=pd.DataFrame(data)
        dftop = dftop.iloc[1: , :]

        #Häufigesten Bauteile mittels Dataframe visualisieren
        with col112:
            st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Häufigsten<br>Komponenten</p>', unsafe_allow_html=True)
            st.write(dftop.astype("object"), width=100)
                  
        st.stop()

#Option Portfolio
if selected_hor == "Portfolio":

    #Überschrift einfügen
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

    #Railway Systems Logo einfügen
    with col2:
        st.image('https://ratek.fi/wp-content/uploads/2020/04/voestalpine_railwaysystems_rgb-color_highres-1024x347.png', width=200)

    #Datei-Uploader einfügen
    with st.expander("Upload"):
        files = st.file_uploader("", accept_multiple_files=True, type=["xlsm"])

    #Sidebar erstellen
    with st.sidebar:
        #VoestAlpine Logo einfügen
        st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/e/ea/Voestalpine_2017_logo.svg/1200px-Voestalpine_2017_logo.svg.png')
        #Sidebar Optionsmenü erstellen
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

    files.sort(key=lambda x: x.name.split("_")[1])

    if files is not None:
        
        try:
            #Selectbox einfügen
            file=st.selectbox("Analyse-File Auswahl", files, format_func=lambda x: x.name.split("_")[0]+x.name.split("_")[1])
            
            df = pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name="Portfolio_Auswertung",
                            na_filter=True)
            
            dfp = pd.read_excel(file,
                            engine="openpyxl",
                            sheet_name="Portfolio",
                            na_filter=True)

            dfp = dfp.astype({'MatNr.':'string'})
            
            if selected_side=="Fertigungsteile":
                dfp=dfp.loc[(dfp['Materialart'] == "FERT") | (dfp['Materialart'] == "60FE")]

            if selected_side=="Zukaufteile":
                dfp=dfp.loc[(dfp['Materialart'] == "HIBE") | (dfp['Materialart'] == "60HI")]
            
            if selected_side=="Normteile":
                dfp=dfp.loc[dfp['Materialart'] == "NORM"]

            col111,col112,col113,col114 = st.columns((1,1,1,1))
            col11, col12 = st.columns((2,5))

            with col112:
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse A</p>', unsafe_allow_html=True)

            with col113:
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse B</p>', unsafe_allow_html=True)
            
            with col114:
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse C</p>', unsafe_allow_html=True)
            
            with col11:
                st. write("")
                st. write("")
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Top Effizienz</p>', unsafe_allow_html=True)
                st.markdown('<p style="text-align: center;color: #FFF;font-size:26px">66 - 100%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Medium Effizienz</p>', unsafe_allow_html=True)
                st.markdown('<p style="text-align: center;color: #FFF;font-size:26px">33 - 66%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")
                st. write("")
                st.markdown('<p style="text-align: center;color: #0082B4;font-size:30px">Low Effizienz</p>', unsafe_allow_html=True)
                st.markdown('<p style="text-align: center;color: #FFF;font-size:26px">0 - 33%</p>', unsafe_allow_html=True)
                st. write("")
                st. write("")
                st. write("")

            with col12:
                option = {
                    "dataset":{},
                    "tooltip":{
                        "show":True,
                    },
                    "series": [
                        {
                        "type": 'liquidFill',
                        "name": "Top-C",
                        "center": ["85%","13%"],
                        "data":[df.at[8,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'orange',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Medium-C",
                        "center": ["85%","42%"],
                        "data":[df.at[7,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'red',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Low-C",
                        "center": ["85%","71%"],
                        "data":[df.at[6,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'red',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Top-B",
                        "center": ["50%","13%"],
                        "data":[df.at[5,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'yellow',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Medium-B",
                        "center": ["50%","42%"],
                        "data":[df.at[4,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'orange',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Low-B",
                        "center": ["50%","71%"],
                        "data":[df.at[3,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'red',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Top-A",
                        "center": ["15%","13%"],
                        "data":[df.at[2,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0eff00',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Medium-A",
                        "center": ["15%","42%"],
                        "data":[df.at[1,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'yellow',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Low-A",
                        "center": ["15%","71%"],
                        "data":[df.at[0,"Anteil"]],
                        "itemStyle": {
                            "color": "#0082B4",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "color": "#FFF",
                        },
                        "radius": "25%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "#0E1117"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": 'orange',
                                "shadowBlur": 0,
                                "shadowColor": '#0082B4'
                        }}},
                        {
                        "type": 'liquidFill',
                        "name": "Reset",
                        "center": ["50%","93%"],
                        "data":[1],
                        "itemStyle": {
                            "color": "white",
                            "shadowBlur": 0
                        },
                        "amplitude": 0,
                        "label": {
                            "formatter": "{a}",
                            "insideColor": "#0082B4",
                            "fontSize":"20"
                        },
                        "radius": "13%",
                        "waveAnimation": 0,
                        "shape": 'roundRect',
                        "backgroundStyle": {
                            "color": "grey"
                        },
                        "outline": {
                            "borderDistance": 5,
                            "itemStyle": {
                                "borderWidth": 5,
                                "borderColor": '#0082B4',
                                "shadowBlur": 0,
                                "shadowColor": 'green'
                        }}}
                        ]
                }
                
                events = {"click": "function(params) { return params.seriesName }"}

                output = st_echarts(option,height=730,events=events)

            if output == "Reset":
                output=None
            
            mask=None

            frame=dfp.loc[dfp['Klasse'] == output]

            frame.dropna(subset=["Materialart"], inplace=True)

            f=frame.astype(str).loc[:, ["MatNr.", "Objektkurztext", "Materialart", "Warengruppe"]]
            
            gb = GridOptionsBuilder.from_dataframe(f)
            grid_options = {
                "defaultColDef": {
                    "minWidth": 5,
                    "editable": False,
                    "filter": True,
                    "resizable": True,
                    "sortable": True
                },
                "suppressFieldDotNotation": True,
                "columnDefs": [
                    {
                    "headerName": "MatNr.",
                    "field": "MatNr.",
                    "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"20px"},
                    "width": 270,
                    "type": []
                    },
                    {
                    "headerName": "Objektkurztext",
                    "field": "Objektkurztext",
                    "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"15px"},
                    "width": 600,
                    "type": []
                    },
                    {
                    "headerName": "Materialart",
                    "field": "Materialart",
                    "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"20px"},
                    "type": []
                    },
                    {
                    "headerName": "Warengruppe",
                    "field": "Warengruppe",
                    "cellStyle": {'background-color': "#0E1117", "font-family":"voestalpine", "font-size":"20px"},
                    "type": []
                    },
                ]
                }
            
            if output is not None:

                if output == "Top-A":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Top-A</p>', unsafe_allow_html=True)

                if output == "Top-B":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Top-B</p>', unsafe_allow_html=True)

                if output == "Top-C":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Top-C</p>', unsafe_allow_html=True)

                if output == "Medium-A":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Medium-A</p>', unsafe_allow_html=True)

                if output == "Medium-B":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Medium-B</p>', unsafe_allow_html=True)

                if output == "Medium-C":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Medium-C</p>', unsafe_allow_html=True)

                if output == "Low-A":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Low-A</p>', unsafe_allow_html=True)

                if output == "Low-B":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Low-B</p>', unsafe_allow_html=True)

                if output == "Low-C":
                    st.markdown('<p style="text-align: center;color: #0082B4;font-size:35px">Klasse Low-C</p>', unsafe_allow_html=True)

            colp1,colp2=st.columns((1,1))

            with colp1:
                sucheM = st.text_input("Materialnummer suchen...",placeholder="z.B. 73100000001A")
            with colp2:
                sucheO = st.text_input("Objektkurztext suchen...",placeholder="z.B. Angriffslappen")

            if output is not None:
                if sucheM and sucheO is not "":
                    mask1 = f["MatNr."].str.contains(sucheM, case=False, na=False)
                    mask2 = f["Objektkurztext"].str.contains(sucheO, case=False, na=False)
                    mask= mask1 & mask2
                
                else:
                    if sucheM is not "":
                        mask = f["MatNr."].str.contains(sucheM, case=False, na=False)
                    if sucheO is not "":
                        mask = f["Objektkurztext"].str.contains(sucheO, case=False, na=False)

                if mask is not None:
                    AgGrid(f[mask],gridOptions=grid_options, theme="streamlit", height= 600,fit_columns_on_grid_load=True, editable=False)
                    
                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        f[mask].to_excel(writer)

                    col1db,col2db,col3db=st.columns((1,1,1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split("_")[0]+file.name.split("_")[1]+"_"+output+".xlsx",
                            mime="application/vnd.ms-excel"
                        )
                else:
                    AgGrid(f,gridOptions=grid_options, theme="streamlit", height= 600,fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        f.to_excel(writer)

                    col1db,col2db,col3db=st.columns((1,1,1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split("_")[0]+file.name.split("_")[1]+"_"+output+".xlsx",
                            mime="application/vnd.ms-excel"
                        )
            else:
                if sucheM and sucheO is not "":
                    mask1 = dfp["MatNr."].str.contains(sucheM, case=False, na=False)
                    mask2 = dfp["Objektkurztext"].str.contains(sucheO, case=False, na=False)
                    mask= mask1 & mask2
                
                else:
                    if sucheM is not "":
                        mask = dfp["MatNr."].str.contains(sucheM, case=False, na=False)
                    if sucheO is not "":
                        mask = dfp["Objektkurztext"].str.contains(sucheO, case=False, na=False)

                if mask is not None:
                    AgGrid(dfp[mask],gridOptions=grid_options, theme="streamlit", height= 600,fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        dfp[mask].to_excel(writer)

                    col1db,col2db,col3db=st.columns((1,1,1))
                    with col2db:
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split("_")[0]+file.name.split("_")[1]+".xlsx",
                            mime="application/vnd.ms-excel"
                        )

                else:
                    AgGrid(dfp,gridOptions=grid_options, theme="streamlit", height= 600,fit_columns_on_grid_load=True, editable=False)

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        dfp.to_excel(writer)

                    col1db,col2db,col3db=st.columns((1,1,1))
                    with col2db:    
                        st.download_button(
                            label="Download Auswertung",
                            data=buffer,
                            file_name=file.name.split("_")[0]+file.name.split("_")[1]+".xlsx",
                            mime="application/vnd.ms-excel"
                        )

            st.stop()

        #Bei fehlenden Input:
        except ValueError:
            st.write("Bitte uploaden Sie die Analyse-Dateien!")
