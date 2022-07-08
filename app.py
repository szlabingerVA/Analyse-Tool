import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
from plotly.subplots import make_subplots

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
    options=("Overview","Effizienz-Analyse","ABC-Analyse"),
    icons=("bezier", "bar-chart-fill", "calculator-fill"),
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
                        "bar":{"color": "green"},
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
        
        #Dataframe anzeigen
        with col13:
            st.dataframe(df.loc[:, ["Komponentennummer", "Effizienz", "Materialart", "Objektsparte"]], height=210)

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
                    "bar":{"color": "green"},
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
            st.dataframe(dftop)
