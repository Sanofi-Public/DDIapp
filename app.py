from shiny import App, Inputs, Outputs, Session, render, ui, reactive
from fpdf import FPDF, HTMLMixin
import pandas as pd 
import io
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font,  Alignment, Color
import asyncio
import os
from shiny.types import ImgData
from datetime import datetime
import html
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors


img_folder = "www"
os.makedirs(img_folder, exist_ok=True)
local_image = os.path.join(img_folder, "Picture.png")  # Replace with your image file

# Glossary dictionary
glossary = {
    "AIV": ["Study code for IND enabling permeability studies", ""],
    "AUC": ["Area Under the plasma Concentration versus time curve extrapolated to infinity", ""],
    "BCRP": ["Breast Cancer Resistance Protein", ""],
    "Cgut": ["Concentration in the intestinal lumen", "Dose/250 mL"],
    "Cmax": ["Maximal systemic concentration in the plasma", ""],
    "Cmax,in": ["Estimated plasma maximal inhibitor concentration at the inlet to the liver", ""],
    "CYP": ["Cytochrome P450", ""],
    "DDI": ["Drug-Drug Interaction", ""],
    "d factor": ["An empirical calibration factor to enable in vitro-to-in vivo induction scaling", "Scaling factor from system used, otherwise assumed 1"],    
    "EC50": ["Concentration of inducer producing 50% of the maximum effect", ""],
    "Emax": ["Maximum fold induction", ""],
    "Fa": ["Fraction of dose absorbed at the gut level", "When Fa value is not available, a default value of 1 is assumed corresponding to entire absorption in gut wall for perpetrator drugs"],
    "Fg": ["Fraction of victim dose escaping metabolism within the walls of the gastrointestinal tract (intestinal wall availability)", "When Fg value is not available, a default value of 1 is assumed corresponding to no metabolism in gut wall for perpetrator drugs"],
    "fm": ["Fraction of hepatic metabolic clearance mediated by enzyme of interest", ""],
    "fu,p": ["Unbound fraction in plasma", "if fu,p < 0.01, keep fu,p=0.01"],
    "fu,mic": ["Unbound fraction in the microsomes, which is used to convert Ki, KI and IC50 values into corresponding unbound parameters", ""],
    "fu,hep": ["Unbound fraction in hepatocytes used to convert EC50 values from induction assays into corresponding unbound parameters", ""],
    "fu,inc": ["Unbound fraction in the assay system (if different from microsomes or hepatocytes)", ""],
    "HLM": ["Human Liver Microsomes", ""],
    "IC50": ["Half maximal inhibitory concentration", ""],
    "ICH": ["International Council for Harmonisation", ""],
    "ICH": ["Study code for IND enabling CYP inhibition studies", ""],
    "IHH": ["Study code for IND enabling CYP induction studies", ""],    
    "INT": ["DDI clinical interaction study", ""],
    "ka": ["Rate constant of absorption (/h) in humans", "If unknown use ka = 6 (worst case scenario)"],
    "Kdeg": ["Endogenous degradation rate constant of an enzyme", ""],
    "Ki": ["Dissociation constant of the enzyme-substrate complex, inhibition constant", ""],
    "KI": ["The concentration of SAR that produces one-half of kinact", ""],
    "kinact": ["Maximal rate of enzyme inactivation", ""],
    "Kobs": ["Observed inactivation rate constant of an affected enzyme", ""],
    "MATE": ["Multidrug And Toxin Extrusion transporter", ""],
    "MBI": ["Mechanism Based Inhibition", ""],
    "MW": ["Molecular Weight", ""],
    "OAT": ["Organic Anion Transporter", ""],
    "OATP": ["Organic Anion Transporting Polypeptide", ""],
    "OCT": ["Organic Cation Transporter", ""],
    "Pgp": ["P-glycoprotein efflux transporter", ""],
    "Qen": ["Intestinal blood flow", "Set by default at 18 L/h"],
    "Qh": ["Hepatic blood flow", "Set by default at 97 L/h"],
    "R": ["Interaction ratio (of the substrate drug with and without the perpetrator)", ""],
    "Rb": ["Blood-to-plasma ratio", "If unknown Rb = 1 for neutrals/bases or Rb = 0.55 for acids/zwitterions"],
    "TDI": ["Time Dependent Inhibition", ""],
    "TRI": ["Study code for IND enabling transporter inhibition studies", ""],
    "UGT": ["UDP-Glucuronosyl-Transferase", ""]
}

# Convert the glossary dictionary to a DataFrame
glossary_df = pd.DataFrame.from_dict(glossary, orient='index', columns=["Definition", "Note"])
glossary_df.reset_index(inplace=True)
glossary_df.rename(columns={'index': 'Term'}, inplace=True)


# Define the UI
app_ui = ui.page_fluid(

     ui.tags.style(
         

            """
        /* Style for Navset Pill List */
        .nav-pills .nav-link.active {
            background-color: #4B0082 !important; /* Dark Purple for Active Tab */
            color: white !important;
            font-size: 25px !important; /* Increase font size */
            padding: 12px 20px !important; /* Increase padding */
        }
        
        .nav-pills .nav-link {
            background-color: #D8BFD8 !important; /* Light Purple for Inactive Tabs */
            color: black !important;
        }


        /* Centering Title */
        .app-title {
            text-align: center;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 20px;
        }
    """,

            """
            .custom-title {
                font-size: 30px;
                color: white;
                background-color: #23004C;
                padding: 20px;
                border-radius: 5px;
                text-align: center;
                width: 100%;
                box-sizing: border-box;
                margin-bottom: 0px;
            }

         .custom-calculate-button {
            background-color: #4CAF50;  /* Green background */
            color: white;               /* White text */
            border: none;               /* Remove border */
            padding: 10px 40px;         /* Add some padding */
            text-align: center;         /* Center text */
            text-decoration: none;      /* Remove underline from the text */
            display: inline-block;      /* Keep it inline */
            font-size: 16px;            /* Increase font size */
            border-radius: 10px;         /* Add rounded corners */
            cursor: pointer;            /* Add pointer on hover */
            width: 200px;
        }
        .custom-calculate-button:hover {
            background-color: #45a049;  /* Darker green on hover */
        }

        .custom-summary-button {
            background-color: #4CAF50;  /* Green background */
            color: white;               /* White text */
            border: none;               /* Remove border */
            padding: 10px 40px;         /* Add some padding */
            text-align: center;         /* Center text */
            text-decoration: none;      /* Remove underline from the text */
            display: inline-block;      /* Keep it inline */
            font-size: 16px;            /* Increase font size */
            border-radius: 10px;         /* Add rounded corners */
            cursor: pointer;            /* Add pointer on hover */
            width: 300px;
        }

        .custom-download-button {
            background-color:rgb(78, 76, 175);  /* Blue background */
            color: white;               /* White text */
            border: none;               /* Remove border */
            padding: 10px 40px;         /* Add some padding */
            text-align: center;         /* Center text */
            text-decoration: none;      /* Remove underline from the text */
            display: inline-block;      /* Keep it inline */
            font-size: 16px;            /* Increase font size */
            border-radius: 10px;         /* Add rounded corners */
            cursor: pointer;            /* Add pointer on hover */
            width: 550px;
        }

        .custum-table {
                        text-align: left;
                        width: 100%;
                        border-collapse: collapse;
        }
        
        .custum-table th {
                        padding: 10px;
                        border: 1px solid #ddd;
                        background-color: #f2f2f2;
                        text-align: left;
        }

        .custum-table td {
                        padding: 10px;
                        border: 1px solid #ddd;
        }   
            """
        ),

        #App Title
        ui.tags.div(
            ui.tags.strong("Static DDI Risk Calculator"), class_="custom-title"
                ),

        ui.row( 

            ui.column(2,
            ui.div(
            ui.div(
            ui.h2(
                    " ", 
             ),

            ui.div(
            
                ),
            ui.input_text("cmp", "Compound", value="SARXXXXXX"),
            ui.input_numeric("mw", "MW (g/mol)", value=300, step=10),
            ui.input_numeric("cmax", "Cmax (ng/mL)", value=500, step=10),
            ui.input_numeric("fup", "fu,p", min=0.01, max=1, value=0.01, step=0.01),
                )
            )
            ),

            ui.column(2,
            ui.div(
            ui.div(
            ui.h2(
                    " ", 
             ),

            ui.div(
            
                ),
            ui.input_numeric("dose", "Dose (mg)", value=200, step=1),
            ui.input_numeric("ka", "ka (/h)", value=6, step=0.1),
            ui.input_numeric("fa", "Fa", value=1, step=0.1, max=1, min=0),
            ui.input_numeric("fg", "Fg", value=1, step=0.1, max=1, min=0), 
                )
            )
            ),

            ui.column(2,
            ui.div(
            ui.div(
            ui.h2(
                    " ", 
             ),

            ui.div(
            
                ),
            # fa, fg, ka, Qh, Rb Input

            ui.input_numeric("qh", "Qh (L/h)", value=97, step=0.1),
            ui.input_numeric("rb", "Rb", value=0.55, step=0.01),
            ui.input_text_area("user_memo", "User memo", value="", height="120px")
                )
            )
            ),

             # Right side: static picture
    ui.column(
        6,
         ui.div(  # Wrapper div to center the image
        ui.output_image("image"),
        style="text-align: center;"  # Centers the image horizontally
    ),
        
    )

            ),
       

    ui.navset_pill_list(
        
      
    
    ui.nav_panel(
        "Enzyme Inhibition",
       
      
    #Enzyme inhibition
    ui.div(
    ui.card(
        ui.card_header("Enzyme inhibition", style="font-size: 24px;"),  # Card Header
            
        # fumic Input
          ui.input_select("fumic", "fu,mic:", {
            "input_fumic": "User input",
            "austin_fumic": "Method of Austin",
            "hallifax_fumic": "Method of Hallifax"
        }),
        ui.output_ui("input_ui_fumic"),
        ui.output_text_verbatim("result_fumic"),

    #Competitive inhibition - hepatic  
    ui.div(
    ui.card(
                ui.h2(
                    "Competitive inhibition - hepatic", 
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseCIH",
                    class_="collapse-toggle collapse-title",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseCIH",
                    style="font-size: 24px; font-weight: bold;"
                ),
                ui.div(
                     
            
    # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Enzyme Selection and Parameters Subsection
                ui.h5("Select Enzymes"),

                ui.row(
                ui.column(6,
                ui.div(
                ui.h5("Mandatory:"),                
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}_{enzyme["substrate"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP1A2", "id":"cyp1a2_ei", "substrate":"Phenacetin"}, 
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2B6", "id":"cyp2b6_ei", "substrate":"Bupropion"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C8", "id":"cyp2c8_ei", "substrate":"Amodiaquine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C9", "id":"cyp2c9_ei", "substrate":"Diclofenac"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C19", "id":"cyp2c19_ei", "substrate":"S-Mephenytoin"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2D6", "id":"cyp2d6_ei", "substrate":"Dextromethorphan"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP3A", "id":"cyp3a_m_ei", "substrate":"Midazolam"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP3A", "id":"cyp3a_t_ei", "substrate":"Testosterone"}
                       ]
                    ]
                )
                )
                ),#ui.column end

                ui.column(6,
                ui.div(
                ui.h5("UGTs:"),
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}_{enzyme["substrate"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A1", "id":"ugt1a1_ei", "substrate":"17β-Estradiol"}, 
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A4", "id":"ugt1a4_ei", "substrate":"Trifluoperazine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A9", "id":"ugt1a9_ei", "substrate":"Propofol"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT2B7", "id":"ugt2b7_ei", "substrate":"Zidovudine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT2B15", "id":"ugt2b15_ei", "substrate":"Oxazepam"}
                       ]
                    ]
                )
                )
                ) # ui.column end
                ), #ui.row end
            )
        )
    ),
    

ui.div(
    ui.input_action_button("calculate_all_ei", "Calculate",class_="custom-calculate-button")  # Calculate Button
),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Competitive inhibition - hepatic)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ei", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{C_{\\text{max}} \\cdot f_{\\text{u,p}}}{MW \\cdot K_{\\text{i}} \\cdot f_{\\text{u,mic}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer ("Red and bold if R >0.02, Green if R =<0.02")
            )
        )
    ), #Competitive inhibition - hepatic card end
     class_="collapse", id="collapseCIH",
    ),
    ),#Competitive inhibition - hepatic card end
    ),#Competitive inhibition - hepatic div end


    ui.div(
    ui.card(
        ui.h2(
             "Irreversible CYP inhibition - hepatic",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseIIH",
                    class_="collapse-toggle collapse-title",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseIIH",
                    style="font-size: 24px; font-weight: bold;"
                ),
    ui.div(
    #ui.card_header("Irreversible CYP inhibition - hepatic", style="font-size: 24px;"),  # Card Header
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Enzyme Selection and Parameters Subsection
                ui.h5("Select Enzymes"),

                ui.row(
                ui.column(6,
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}_{enzyme["substrate"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP1A2", "id":"cyp1a2_tdi", "substrate":"Phenacetin", "kdeg":0.0180}, 
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2B6", "id":"cyp2b6_tdi", "substrate":"Bupropion", "kdeg":0.0216},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C8", "id":"cyp2c8_tdi", "substrate":"Amodiaquine", "kdeg":0.0318},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C9", "id":"cyp2c9_tdi", "substrate":"Diclofenac", "kdeg":0.0066},
                      ]
                    ]
                )
                ),#ui.column end

                ui.column(6,
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}_{enzyme["substrate"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C19", "id":"cyp2c19_tdi", "substrate":"S-Mephenytoin", "kdeg":0.0264},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2D6", "id":"cyp2d6_tdi", "substrate":"Dextromethorphan", "kdeg":0.0138},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP3A", "id":"cyp3a_m_tdi", "substrate":"Midazolam", "kdeg":0.0192},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP3A", "id":"cyp3a_t_tdi", "substrate":"Testosterone", "kdeg":0.0192},
                   ]
                    ]
                )
                )#ui.column end
                )#ui.row end
            ),#parameters end
        )
    ),

    # Calculate Card
    ui.div(
            ui.input_action_button("calculate_all_tdi", "Calculate",class_="custom-calculate-button")  # Calculate Button    
        ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Irreversible CYP inhibition - hepatic)"),  # Card Header
            ui.card_body(
               ui.output_table("result_table_tdi", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{K_{\\text{obs}} + K_{\\text{deg}}}{K_{\\text{deg}}}, K_{\\text{obs}} = \\frac{K_{\\text{inact}} \\cdot 5 \\cdot \\frac{C_{\\text{max}}}{MW} \\cdot f_{\\text{u,p}}}{K_{\\text{I}} \\cdot f_{\\text{u,mic}} + 5 \\cdot \\frac{C_{\\text{max}}}{MW} \\cdot f_{\\text{u,p}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >1.25, Green if R =<1.25")
            )
        )
    ),
    class_="collapse", id="collapseIIH",
    ),
    ),#Irreversible CYP inhibition - hepatic card end
    ),#Irreversible CYP inhibition - hepatic div end


    ui.div(
    ui.card(
           ui.h2(
             "Competitive CYP inhibition - intestinal",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseCII",
                    class_="collapse-toggle collapse-title",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseCII",
                    style="font-size: 24px; font-weight: bold;"
                ),
    ui.div(
        
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Enzyme Selection and Parameters Subsection
                ui.h5("Select Enzymes"),             
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}_{enzyme["substrate"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Competitive CYP inhibition - intestinal","name":"CYP3A", "id":"cyp3a_m_ei_int", "substrate":"Midazolam"},
                            {"attribute":"Competitive CYP inhibition - intestinal","name":"CYP3A", "id":"cyp3a_t_ei_int", "substrate":"Testosterone"},
                       ]
                    ]
                ),
            )
        )
    ),

    # Calculate Card
    ui.div(
     ui.input_action_button("calculate_all_ei_int", "Calculate",class_="custom-calculate-button")  # Calculate Button         
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Competitive CYP inhibition - intestinal)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ei_int", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{Dose \\cdot 1000000}{MW \\cdot 250 \\cdot K_{\\text{i}} \\cdot f_{\\text{u,mic}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >10, Green if R =<10")
            )
        )
    ),
    class_="collapse", id="collapseCII",
    ), 
    ),#Competitive CYP inhibition - intestinal card end
    )#Competitive CYP inhibition - intestinal div end

    )#Enzyme inhibition card end
    )#Enzyme inhibition div end
    ),#Enzyme inhibition panel end


    ui.nav_panel(
        "Enzyme Induction",
    # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Enzyme Induction", style="font-size: 24px;"),  # Card Header
            ui.card_body(
                # d_factor Input
                ui.input_numeric("d_factor", "d factor", value=1, step=0.1),
                
                # fuhep Input
                ui.input_select("fuhep", "fu,hep:", {
                "input_fuhep": "User input",
                "austin_fuhep": "Method of Austin",
                "kilfold_fuhep": "Method of Kilfold"
                }),
                ui.output_ui("input_ui_fuhep"),
                ui.output_text_verbatim("result_fuhep"),
                
                # fuhep Input
                # ui.input_numeric("fuhep", "fu.hep", value=0.1, step=0.01),

                # Enzyme Selection and Parameters Subsection
                ui.h5("Select Enzymes:"),
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Enzyme induction","name":"CYP1A2", "id":"cyp1a2_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2B6", "id":"cyp2b6_ed"},
                            {"attribute":"Enzyme induction","name":"CYP3A4", "id":"cyp3a4_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2C8", "id":"cyp2c8_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2C9", "id":"cyp2c9_ed"}, 
                            {"attribute":"Enzyme induction","name":"CYP2C19", "id":"cyp2c19_ed"}, 
                            {"attribute":"Enzyme induction","name":"UGT1A1", "id":"ugt1a1_ed"}, 
                            {"attribute":"Enzyme induction","name":"UGT1A4", "id":"ugt1a4_ed"}                          
                        ]
                    ]
                )
            )
        )
    ),

    # Calculate Card
    ui.div(   
        ui.input_action_button("calculate_all_ed", "Calculate",class_="custom-calculate-button")  # Calculate Button         
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Enzyme induction)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ed", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{1}{1 + \\frac{d \\cdot E_{\\text{max}} \\cdot 10 \\cdot \\frac{C_{\\text{max}}}{MW} \\cdot f_{\\text{u,p}}}{EC_{\\text{50}} \\cdot f_{\\text{u,hep}} + 10 \\cdot \\frac{C_{\\text{max}}}{MW} \\cdot f_{\\text{u,p}}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red if R <0.2, Orange if 0.2=< R <0.5, Blue if 0.5=< R <0.8, Green if R >= 0.8") 
            )
        )
    )
),# Enzyme induction panel end




    ui.nav_panel(
        "Net effect",
    # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Net effect", style="font-size: 24px;"),  # Card Header
            ui.card_body(
                
                ui.input_numeric("qen", "Qen (L/h)", value=18, step=0.1),
                
                # Enzyme Selection and Parameters Subsection
                ui.h5("Select Enzymes:"),
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{enzyme["id"]}', f'{enzyme["name"]}', value=False),
                            ui.output_ui(f'parameters_{enzyme["id"]}'),
                            class_="enzyme-block"
                        )
                        for enzyme in [
                            {"attribute":"Net effect","name":"CYP1A2", "id":"cyp1a2_ne", "substrate":"Phenacetin"},
                            {"attribute":"Net effect","name":"CYP2B6", "id":"cyp2b6_ne", "substrate":"Bupropion"},
                            {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_ne"},
                            # {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_t_ne", "substrate":"Testosterone"},
                            # {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_m_ne", "substrate":"Midazolam"},
                            {"attribute":"Net effect","name":"CYP2C8", "id":"cyp2c8_ne", "substrate":"Amodiaquine"},
                            {"attribute":"Net effect","name":"CYP2C9", "id":"cyp2c9_ne", "substrate":"Diclofenac"}, 
                            {"attribute":"Net effect","name":"CYP2C19", "id":"cyp2c19_ne", "substrate":"S-Mephenytoin"},                         
                            {"attribute":"Net effect","name":"CYP2D6", "id":"cyp2d6_ne", "substrate":"Dextromethorphan"},                         
                            {"attribute":"Net effect","name":"UGT1A1", "id":"ugt1a1_ne", "substrate":"17β-Estradiol"},                         
                            {"attribute":"Net effect","name":"UGT1A4", "id":"ugt1a4_ne", "substrate":"Trifluoperazine"}                      
                        ]
                    ]
                )
            )
        )
    ),

    # Calculate Card
    ui.div(   
        ui.input_action_button("calculate_all_ne", "Calculate",class_="custom-calculate-button")  # Calculate Button         
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Net effect)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ne", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$AUCR (R) = \\frac{1}{[A_{\\text{g}} \\cdot B_{\\text{g}} \\cdot C_{\\text{g}}] \\cdot (1 - F_{\\text{g}}) + F_{\\text{g}}} \\cdot \\frac{1}{[A_{\\text{h}} \\cdot B_{\\text{h}} \\cdot C_{\\text{h}}] \\cdot f_{\\text{m}} + (1 - f_{\\text{m}})}$$",
                    "$$A_{\\text{g}} = \\frac{1}{1 + \\frac{[I]_{\\text{g}}}{K_{\\text{i}} \\cdot f_{\\text{u,mic}}}}, B_{\\text{g}} = 1, C_{\\text{g}} = 1$$",
                    "$$A_{\\text{h}} = \\frac{1}{1 + \\frac{[I]_{\\text{h}}}{K_{\\text{i}} \\cdot f_{\\text{u,mic}}}}, B_{\\text{h}} = \\frac{K_{\\text{deg,h}}}{K_{\\text{deg,h}} + \\frac{[I]_{\\text{h}} \\cdot K_{\\text{inact}}}{[I]_{\\text{h}} + K_{\\text{I}} \\cdot f_{\\text{u,mic}}}}, C_{\\text{h}} = \\frac{d \\cdot E_{\\text{max}} \\cdot [I]_{\\text{h}}}{[I]_{\\text{h}} + EC_{\\text{50}} \\cdot f_{\\text{u,hep}}}$$",
                    "$$[I]_{\\text{g}} = F_{\\text{a}} \\cdot k_{\\text{a}} \\cdot \\frac{Dose \\cdot 1000}{MW \\cdot Q_{\\text{en}}}, [I]_{\\text{h}} = f_{\\text{u,p}} \\cdot (\\frac{C_{\\text{max}}}{MW} + \\frac{F_{\\text{a}} \\cdot F_{\\text{g}} \\cdot k_{\\text{a}} \\cdot Dose \\cdot 1000}{MW \\cdot Q_{\\text{h}} \\cdot R_{\\text{b}}})$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red if R <0.8 or R > 1.25, Green if 0.8 =< R =< 1.25"), 
                ui.tags.footer("If there are no input values for each enzyme from 'Enzyme inhibition' or 'Enzyme induction' panels, Ag, Bg, Cg, Ah, Bh or Ch are set to 1") 
            )
        )
    )
),# net effect panel end




ui.nav_panel(
    "Transporter inhibition",

    #Transporter inhibition
    ui.div(
    ui.card(
        ui.card_header("Transporter inhibition", style="font-size: 24px;"),  # Card Header
        ui.input_numeric("fuinc", "fu,inc", value=0.1, step=0.01),
    ui.div(
    ui.card(
        ui.h2(
            "Transporter inhibition - intestinal efflux",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseTIIE",
                    class_="collapse-toggle collapse-title",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseTIIE",
                    style="font-size: 24px; font-weight: bold;"
                ),

            
            ui.div(
       
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Transporter Selection and Parameters Subsection
                ui.h5("Select Transporters"),             
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{transporter["id"]}', f'{transporter["name"]}', value=False),
                            ui.output_ui(f'parameters_{transporter["id"]}'),
                            class_="transporter-block"
                        )
                        for transporter in [
                            {"attribute":"Transporter inhibition - intestinal efflux","name":"Pgp", "id":"pgp_ti_ie"},
                            {"attribute":"Transporter inhibition - intestinal efflux","name":"BCRP", "id":"bcrp_ti_ie"},
                            ]
                    ]
                ),
            )
        )
    ),

    # Calculate Card
    ui.div(
        ui.input_action_button("calculate_all_ti_ie", "Calculate",class_="custom-calculate-button")  # Calculate Button
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Transporter inhibition - intestinal efflux)"),  # Card Header
            ui.card_body(
               ui.output_table("result_table_ti_ie", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{Dose \\cdot 1000000}{MW \\cdot 250 \\cdot (IC_{\\text{50}} or K_{\\text{i}}) \\cdot f_{\\text{u,inc}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >10, Green if R =<10")
            )
        )
    ),
    class_="collapse", id="collapseTIIE",
    ),
    ),#Transporter inhibition - intestinal efflux card end
    ),#Transporter inhibition - intestinal efflux div end


    ui.div(
    ui.card(
        ui.h2(
            "Transporter inhibition - hepatic uptake",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseTIHU",
                    class_="collapse-toggle collapse-title",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseTIHU",
                    style="font-size: 24px; font-weight: bold;"
                ),

            
            ui.div(
        
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Transporter Selection and Parameters Subsection
                ui.h5("Select Transporters"),             
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{transporter["id"]}', f'{transporter["name"]}', value=False),
                            ui.output_ui(f'parameters_{transporter["id"]}'),
                            class_="transporter-block"
                        )
                        for transporter in [
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OATP1B1", "id":"oatp1b1_ti_hu"},
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OATP1B3", "id":"oatp1b3_ti_hu"},
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OCT1", "id":"oct1_ti_hu"},
                            ]
                    ]
                ),
            )
        )
    ),

    # Calculate Card
    ui.div(
       ui.input_action_button("calculate_all_ti_hu", "Calculate",class_="custom-calculate-button")  # Calculate Button
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Transporter inhibition - hepatic uptake)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ti_hu", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),    
                ui.tags.div(
                    "$$R = \\frac{C_{\\text{max,in}} \\cdot f_{\\text{u,p}}}{(IC_{\\text{50}} or K_{\\text{i}}) \\cdot f_{\\text{u,inc}}}, C_{\\text{max,in}} = \\frac{C_{\\text{max}}}{MW} + \\frac{F_{\\text{a}} \\cdot F_{\\text{g}} \\cdot k_{\\text{a}} \\cdot Dose \\cdot 1000}{MW \\cdot Q_{\\text{h}} \\cdot R_{\\text{b}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >0.1, Green if R =<0.1")
            )
        )
    ),
        class_="collapse", id="collapseTIHU",
    ),
    ),#Transporter inhibition - hepatic uptake card end
    ),#Transporter inhibition - hepatic uptake div end


    ui.div(
    ui.card(
        ui.h2(
            "Transporter inhibition - renal uptake",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseTIRU",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseTIRU",
                    style="font-size: 24px; font-weight: bold;"
                ),

            
            ui.div(
       
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Transporter Selection and Parameters Subsection
                ui.h5("Select Transporters"),             
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{transporter["id"]}', f'{transporter["name"]}', value=False),
                            ui.output_ui(f'parameters_{transporter["id"]}'),
                            class_="transporter-block"
                        )
                        for transporter in [
                            {"attribute":"Transporter inhibition - renal uptake","name":"OAT1", "id":"oatp1b1_ti_ru"},
                            {"attribute":"Transporter inhibition - renal uptake","name":"OAT3", "id":"oatp1b3_ti_ru"},
                            {"attribute":"Transporter inhibition - renal uptake","name":"OCT2", "id":"oct1_ti_ru"},
                            ]
                    ]
                ),
            )
        )
    ),
    # Calculate Card
    ui.div(
         ui.input_action_button("calculate_all_ti_ru", "Calculate",class_="custom-calculate-button")  # Calculate Button   
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Transporter inhibition - renal uptake)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ti_ru", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{C_{\\text{max}} \\cdot f_{\\text{u,p}}}{MW \\cdot (IC_{\\text{50}} or K_{\\text{i}}) \\cdot f_{\\text{u,inc}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >0.1, Green if R =<0.1")
            )
        )
    ),
     class_="collapse", id="collapseTIRU",
    ),
    ),#Transporter inhibition - renal uptake card end
    ),#Transporter inhibition - renal uptake div end


    ui.div(
    ui.card(
        ui.h2(
            "Transporter inhibition - renal efflux",
                    ui.span("+", class_="collapseSymbol", style="float: right;"),
                    href="#collapseTIRE",
                    data_bs_toggle="collapse", 
                    role="button", 
                    aria_expanded="false", 
                    aria_controls="collapseTIRE",
                    style="font-size: 24px; font-weight: bold;"
                ),

            
            ui.div(
        # Card Header
   # Parameters Card
    ui.div(
        ui.card(
            ui.card_header("Parameters"),  # Card Header         
            ui.card_body(
                # Transporter Selection and Parameters Subsection
                ui.h5("Select Transporters"),             
                ui.div(
                    *[
                        ui.div(
                            ui.input_checkbox(f'select_{transporter["id"]}', f'{transporter["name"]}', value=False),
                            ui.output_ui(f'parameters_{transporter["id"]}'),
                            class_="transporter-block"
                        )
                        for transporter in [
                            {"attribute":"Transporter inhibition - renal efflux","name":"MATE1", "id":"mate1_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"MATE2K", "id":"mate2k_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"Pgp", "id":"pgp_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"BCRP", "id":"bcrp_ti_re"},                           
                             ]
                    ]
                ),
            )
        )
    ),
    # Calculate Card
    ui.div(
        ui.input_action_button("calculate_all_ti_re", "Calculate",class_="custom-calculate-button")  # Calculate Button
    ),

    # Results Card
    ui.div(
        ui.card(
            ui.card_header("Results (Transporter inhibition - renal efflux)"),  # Card Header
            ui.card_body(
                ui.output_table("result_table_ti_re", class_="custum-table"),  # Display results in a table

                    # Display LaTeX Equation
                ui.tags.head(
                    ui.tags.script(src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js")
                    ),
                ui.tags.div(
                    "$$R = \\frac{C_{\\text{max}} \\cdot f_{\\text{u,p}}}{MW \\cdot (IC_{\\text{50}} or K_{\\text{i}}) \\cdot f_{\\text{u,inc}}}$$",
                    style="text-align: center; font-size: 1.2em; margin-top: 20px;"
                ),
                ui.tags.footer("Red and bold if R >0.02, Green if R =<0.02")
            )
        )
    ),
     class_="collapse", id="collapseTIRE",
    ),
    ),#Transporter inhibition - renal efflux card end
    ),#Transporter inhibition - renal efflux div end

    )#Transporter inhibition card end
    )#Transporter inhibition div end
), # Transporter inhibition panel end

ui.nav_panel(
    "Summary",

    ui.div(
        ui.card(
            ui.card_header("Summary", style="font-size: 24px;"),  # Card Header
            ui.card_body(
                # Action buttons
                ui.div(
                    ui.input_action_button("summary", "Create summary tables",class_="custom-summary-button")  # Summary Button
                ),
                ui.div(
                    ui.download_button("handle_download_xlsx", "Download excel file (Click after creating summary tables)",class_="custom-download-button")   # Download Button
                ),

                ui.div(
                    ui.download_button("handle_download_pdf", "Download PDF file (Click after creating summary tables)",class_="custom-download-button")   # Download Button
                ),

                # Results Card
                ui.div(
                    ui.output_table("summary_table", class_="custum-table"),  # Display results in a table
                )
 
            )#ui.card.body summary end
        )#ui.card summary end
    ),
), # summary panel end

ui.nav_panel(
    "Glossary",

    ui.div(
        ui.card(
            ui.card_header("Glossary", style="font-size: 24px;"),  # Card Header
            ui.card_body(
                ui.tags.p("Terms and Definitions", style="font-size: 20px; font-weight: bold; margin-top: 20px;"), 
                ui.output_table("glossary_table", class_="custum-table"),  
                ui.tags.p("Useful Links", style="font-size: 20px; font-weight: bold; margin-top: 20px;"), 
                ui.tags.a("M12 Drug Interaction Studies (US FDA, Aug 2024)", href="https://www.fda.gov/media/161199/download", target="_blank", style="display: block; margin-top: 10px;"),
                ui.tags.a("M12 Drug Interaction Studies: Q&A (US FDA, Aug 2024)", href="https://www.fda.gov/media/180488/download", target="_blank", style="display: block; margin-top: 10px;"),
                ui.tags.a("ICH M12 Guideline on drug interaction studies - Scientific guideline (EMA, HP link)", href="https://www.ema.europa.eu/en/ich-m12-drug-interaction-studies-scientific-guideline", target="_blank", style="display: block; margin-top: 10px;"),
                ui.tags.a("ICH M12 薬物相互作用試験 (PMDA, HP link)", href="https://www.pmda.go.jp/int-activities/int-harmony/ich/0101.html", target="_blank", style="display: block; margin-top: 10px;"),
                ui.tags.a("CDE公开征求ICH《M12：药物相互作用》指导原则及问答文件实施建议和中文版意见 (NMPA, HP link)", href="https://www.cnpharm.com/c/2024-07-18/1050390.shtml", target="_blank", style="display: block; margin-top: 10px;"),
            )#ui.card.body glossary end
        )#ui.card glossary end
    ),
), # glossary panel end

),# ui.navset_pill_list end


# Go to top button
ui.div(
    ui.input_action_button("top_button", "Go to Top"),
    style="position: fixed; bottom: 10px; right: 10px;"
),
ui.tags.script("""
    document.getElementById('top_button').onclick = function() {
        window.scrollTo({top: 0, behavior: 'smooth'});
    };
""")

)# app ui end


# Define the server logic
def server(input: Inputs, output: Outputs, session: Session):

    @render.image
    def image():
        from pathlib import Path

        dir = Path(__file__).resolve().parent

        img: ImgData = {"src": str(dir / "DDI_logo.png"), "width": "425px"}
        #img: ImgData = {"src": str(dir / "DDI_logo2.png"), "width": "425px"}
        #img: ImgData = {"src": str(dir / "ddi-app-logo2.png"), "width": "480px"}

        return img

    # List of all enzymes
    all_enzymes_ei = [ 
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP1A2", "id":"cyp1a2_ei", "substrate":"Phenacetin"}, 
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2B6", "id":"cyp2b6_ei", "substrate":"Bupropion"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C8", "id":"cyp2c8_ei", "substrate":"Amodiaquine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C9", "id":"cyp2c9_ei", "substrate":"Diclofenac"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2C19", "id":"cyp2c19_ei", "substrate":"S-Mephenytoin"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP2D6", "id":"cyp2d6_ei", "substrate":"Dextromethorphan"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP3A", "id":"cyp3a_m_ei", "substrate":"Midazolam"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"CYP3A", "id":"cyp3a_t_ei", "substrate":"Testosterone"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A1", "id":"ugt1a1_ei", "substrate":"17β-Estradiol"}, 
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A4", "id":"ugt1a4_ei", "substrate":"Trifluoperazine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT1A9", "id":"ugt1a9_ei", "substrate":"Propofol"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT2B7", "id":"ugt2b7_ei", "substrate":"Zidovudine"},
                            {"attribute":"Competitive enzyme inhibition - hepatic","name":"UGT2B15", "id":"ugt2b15_ei", "substrate":"Oxazepam"}
    ]       

    all_enzymes_tdi = [
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP1A2", "id":"cyp1a2_tdi", "substrate":"Phenacetin", "kdeg":0.0180}, 
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2B6", "id":"cyp2b6_tdi", "substrate":"Bupropion", "kdeg":0.0216},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C8", "id":"cyp2c8_tdi", "substrate":"Amodiaquine", "kdeg":0.0318},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C9", "id":"cyp2c9_tdi", "substrate":"Diclofenac", "kdeg":0.0066},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2C19", "id":"cyp2c19_tdi", "substrate":"S-Mephenytoin", "kdeg":0.0264},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP2D6", "id":"cyp2d6_tdi", "substrate":"Dextromethorphan", "kdeg":0.0138},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP3A", "id":"cyp3a_m_tdi", "substrate":"Midazolam", "kdeg":0.0192},
                            {"attribute":"Irreversible CYP inhibition - hepatic","name":"CYP3A", "id":"cyp3a_t_tdi", "substrate":"Testosterone", "kdeg":0.0192},
    ]

    all_enzymes_ei_int = [
                            {"attribute":"Competitive CYP inhibition - intestinal","name":"CYP3A", "id":"cyp3a_m_ei_int", "substrate":"Midazolam"},
                            {"attribute":"Competitive CYP inhibition - intestinal","name":"CYP3A", "id":"cyp3a_t_ei_int", "substrate":"Testosterone"},
    ]

    all_enzymes_ed = [
                            {"attribute":"Enzyme induction","name":"CYP1A2", "id":"cyp1a2_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2B6", "id":"cyp2b6_ed"},
                            {"attribute":"Enzyme induction","name":"CYP3A4", "id":"cyp3a4_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2C8", "id":"cyp2c8_ed"},
                            {"attribute":"Enzyme induction","name":"CYP2C9", "id":"cyp2c9_ed"}, 
                            {"attribute":"Enzyme induction","name":"CYP2C19", "id":"cyp2c19_ed"}, 
                            {"attribute":"Enzyme induction","name":"UGT1A1", "id":"ugt1a1_ed"}, 
                            {"attribute":"Enzyme induction","name":"UGT1A4", "id":"ugt1a4_ed"}
    ]

    all_enzymes_ne = [
                            {"attribute":"Net effect","name":"CYP1A2", "id":"cyp1a2_ne", "substrate":"Phenacetin"},
                            {"attribute":"Net effect","name":"CYP2B6", "id":"cyp2b6_ne", "substrate":"Bupropion"},
                            # {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_t_ne", "substrate":"Testosterone"},
                            # {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_m_ne", "substrate":"Midazolam"},
                            {"attribute":"Net effect","name":"CYP2C8", "id":"cyp2c8_ne", "substrate":"Amodiaquine"},
                            {"attribute":"Net effect","name":"CYP2C9", "id":"cyp2c9_ne", "substrate":"Diclofenac"}, 
                            {"attribute":"Net effect","name":"CYP2C19", "id":"cyp2c19_ne", "substrate":"S-Mephenytoin"},                         
                            {"attribute":"Net effect","name":"CYP2D6", "id":"cyp2d6_ne", "substrate":"Dextromethorphan"},                         
                            {"attribute":"Net effect","name":"UGT1A1", "id":"ugt1a1_ne", "substrate":"17β-Estradiol"},                         
                            {"attribute":"Net effect","name":"UGT1A4", "id":"ugt1a4_ne", "substrate":"Trifluoperazine"}   
    ]

    all_enzymes_ne_cyp3a4 = [
                            {"attribute":"Net effect","name":"CYP3A4", "id":"cyp3a4_ne"},
    ]    

    cyp3a_substrates = [
                            {"substrate":"Midazolam"},
                            {"substrate":"Testosterone"}
    ]


    all_transporters_ie = [
                            {"attribute":"Transporter inhibition - intestinal efflux","name":"Pgp", "id":"pgp_ti_ie"},
                            {"attribute":"Transporter inhibition - intestinal efflux","name":"BCRP", "id":"bcrp_ti_ie"},
    ]

    all_transporters_hu = [
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OATP1B1", "id":"oatp1b1_ti_hu"},
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OATP1B3", "id":"oatp1b3_ti_hu"},
                            {"attribute":"Transporter inhibition - hepatic uptake","name":"OCT1", "id":"oct1_ti_hu"},
    ]

    all_transporters_ru = [
                            {"attribute":"Transporter inhibition - renal uptake","name":"OAT1", "id":"oatp1b1_ti_ru"},
                            {"attribute":"Transporter inhibition - renal uptake","name":"OAT3", "id":"oatp1b3_ti_ru"},
                            {"attribute":"Transporter inhibition - renal uptake","name":"OCT2", "id":"oct1_ti_ru"},
    ]

    all_transporters_re = [
                            {"attribute":"Transporter inhibition - renal efflux","name":"MATE1", "id":"mate1_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"MATE2K", "id":"mate2k_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"Pgp", "id":"pgp_ti_re"},
                            {"attribute":"Transporter inhibition - renal efflux","name":"BCRP", "id":"bcrp_ti_re"},                           
    ]


    # fu,mic
    @output
    @render.ui
    def input_ui_fumic():
        if input.fumic() == "input_fumic":
            return ui.input_numeric("input_value_fumic", "fu,mic - experimental value", value=0.1, step=0.01)
        elif input.fumic() == "austin_fumic":
            return ui.TagList(
                ui.input_numeric("log_p_d_fumic", "Log P/D: log P for pKa >7.4 and log D7.4 for pKa =<7.4", value=1, step=0.1),
                ui.input_numeric("c_value", "C: microsomal protein concentration (g/L)", value=0.1, step=0.01),
                ui.tags.footer("fu,mic = 1 / (1 + C * 10**(0.56 * log P/D - 1.41))")
                    )
        elif input.fumic() == "hallifax_fumic":
            return ui.TagList(
                ui.input_numeric("log_p_d_fumic", "Log P/D: log P for pKa >7.4 and log D7.4 for pKa =<7.4", value=1, step=0.1),
                ui.input_numeric("c_value", "C: microsomal protein concentration (g/L)", value=0.1, step=0.01),
                ui.tags.footer("fu,mic = 1 / (1 + C * 10**(0.072 * log P/D **2 + 0.067 * log P/D - 1.126))")
            )

    @reactive.Calc
    def calculate_fumic():
        if input.fumic() == "input_fumic":
            return float(input.input_value_fumic())
        elif input.fumic() == "austin_fumic":
            log_p_d_fumic = input.log_p_d_fumic()
            c_value = input.c_value()
            return 1 / (1 + c_value * 10**(0.56 * log_p_d_fumic - 1.41))
        elif input.fumic() == "hallifax_fumic":
            log_p_d_fumic = input.log_p_d_fumic()
            c_value = input.c_value()
            return 1 / (1 + c_value * 10**(0.072 * log_p_d_fumic**2 + 0.067 * log_p_d_fumic - 1.126))
    @output
    @render.text
    def result_fumic():
        global fumic
        fumic = calculate_fumic()
        return f"fu,mic: {round(fumic,3)}"


    # fu,hep
    @output
    @render.ui
    def input_ui_fuhep():
        if input.fuhep() == "input_fuhep":
            return ui.input_numeric("input_value_fuhep", "fu,hep - experimental value", value=0.1, step=0.01)
        elif input.fuhep() == "austin_fuhep":
            return ui.TagList(
                ui.input_numeric("log_p_d_fuhep", "Log P/D: log P for pKa >7.4 and log D7.4 for pKa =<7.4", value=1, step=0.1),
                ui.tags.footer("fu,hep = 1 / (1 + 10**(0.40 * log P/D - 1.38))")
            )
        elif input.fuhep() == "kilfold_fuhep":
            return ui.TagList(
                ui.input_numeric("log_p_d_fuhep", "Log P/D: log P for pKa >7.4 and log D7.4 for pKa =<7.4", value=1, step=0.1),
                ui.tags.footer("fu,hep = 1 / (1 + 0.625 * 10**(0.072 * log P/D**2 + 0.067 * log P/D - 1.126))")
            )

    @reactive.Calc
    def calculate_fuhep():
        if input.fuhep() == "input_fuhep":
            return float(input.input_value_fuhep())
        elif input.fuhep() == "austin_fuhep":
            log_p_d_fuhep = input.log_p_d_fuhep()
            return 1 / (1 + 10**(0.40 * log_p_d_fuhep - 1.38))
        elif input.fuhep() == "kilfold_fuhep":
            log_p_d_fuhep = input.log_p_d_fuhep()
            return 1 / (1 + 0.625 * 10**(0.072 * log_p_d_fuhep**2 + 0.067 * log_p_d_fuhep - 1.126))

    @output
    @render.text
    def result_fuhep():
        global fuhep
        fuhep = calculate_fuhep()
        return f"fu,hep: {round(fuhep,3)}"



    # Dynamically render parameters for each enzyme and transporter 
    def make_render_parameters_ei(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ki_{enzyme['id']}", f"Ki ({enzyme['name']}_{enzyme['substrate']}) (μmol/L)", value=10, step=0.1),
                )
            else:
                return None

    def make_render_parameters_tdi(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ki_tdi_{enzyme['id']}", f"KI ({enzyme['name']}_{enzyme['substrate']}) (μmol/L)", value=10, step=0.1),
                    ui.input_numeric(f"kinact_{enzyme['id']}", f"Kinact ({enzyme['name']}_{enzyme['substrate']}) (/h)", value=0.1, step=0.01),                   
                    ui.input_numeric(f"kdeg_{enzyme['id']}", f"Kdeg ({enzyme['name']}_{enzyme['substrate']}) (/h)", value=enzyme['kdeg'], step=0.0001),                   
                )
            else:
                return None

    def make_render_parameters_ei_int(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ki_{enzyme['id']}", f"Ki ({enzyme['name']}_{enzyme['substrate']}) (μmol/L)", value=10, step=0.1)                  
                )
            else:
                return None

    def make_render_parameters_ed(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"emax_{enzyme['id']}", f"Emax ({enzyme['name']})", value=0.3, step=0.1),
                    ui.input_numeric(f"ec50_{enzyme['id']}", f"EC50 ({enzyme['name']}) (μmol/L)", value=20, step=1),
                )
            else:
                return None

    def make_render_parameters_ne(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"fm_{enzyme['id']}", f"fm,{enzyme['name']}", value=0.3, step=0.01, max=1, min=0),
                )                    
            else:
                return None

    def make_render_parameters_ne_cyp3a4(enzyme):
        @output(id=f"parameters_{enzyme['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{enzyme['id']}"]():
                return ui.div(
                    ui.input_numeric(f"fm_{enzyme['id']}", f"fm,{enzyme['name']}", value=0.3, step=0.01, max=1, min=0),
                    ui.div(
                        *[ui.input_checkbox(f'select_{enzyme["id"]}_{substrate["substrate"]}', substrate["substrate"], value=False) for substrate in cyp3a_substrates],
                        style="margin-left: 50px;"  
                    )
                )
            else:
                return None

    def make_render_parameters_ti_ie(transporter):
        @output(id=f"parameters_{transporter['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{transporter['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ic50_{transporter['id']}", f"IC50 or Ki ({transporter['name']}) (μmol/L)", value=10, step=0.1),
                )
            else:
                return None

    def make_render_parameters_ti_hu(transporter):
        @output(id=f"parameters_{transporter['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{transporter['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ic50_{transporter['id']}", f"IC50 or Ki ({transporter['name']}) (μmol/L)", value=10, step=0.1),
                )
            else:
                return None

    def make_render_parameters_ti_ru(transporter):
        @output(id=f"parameters_{transporter['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{transporter['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ic50_{transporter['id']}", f"IC50 or Ki ({transporter['name']}) (μmol/L)", value=10, step=0.1),
                )
            else:
                return None

    def make_render_parameters_ti_re(transporter):
        @output(id=f"parameters_{transporter['id']}")
        @render.ui
        def render_parameters():
            if input[f"select_{transporter['id']}"]():
                return ui.div(
                    ui.input_numeric(f"ic50_{transporter['id']}", f"IC50 or Ki ({transporter['name']}) (μmol/L)", value=10, step=0.1),
                )
            else:
                return None

    # Register render functions for each enzyme
    for enzyme in all_enzymes_ei:
        make_render_parameters_ei(enzyme)

    for enzyme in all_enzymes_tdi:
        make_render_parameters_tdi(enzyme)

    for enzyme in all_enzymes_ei_int:
        make_render_parameters_ei_int(enzyme)

    for enzyme in all_enzymes_ed:
        make_render_parameters_ed(enzyme)

    for enzyme in all_enzymes_ne:
        make_render_parameters_ne(enzyme)

    for enzyme in all_enzymes_ne_cyp3a4:
        make_render_parameters_ne_cyp3a4(enzyme)

    for transporter in all_transporters_ie:
        make_render_parameters_ti_ie(transporter)

    for transporter in all_transporters_hu:
        make_render_parameters_ti_hu(transporter)

    for transporter in all_transporters_ru:
        make_render_parameters_ti_ru(transporter)

    for transporter in all_transporters_re:
        make_render_parameters_ti_re(transporter)

    # Calculate results for selected enzymes and transporters

    #Competitive inhibition - hepatic
    @output
    @render.ui
    @reactive.event(input.calculate_all_ei)
    def result_table_ei():
        # Collect selected enzymes
        selected_enzymes = [enzyme for enzyme in all_enzymes_ei if input[f"select_{enzyme['id']}"]()]
        global results_df_ei 
        global results_df_ei_net
        results_ei = []
        results_ei_net = []

        # Perform calculations for each selected enzyme
        for enzyme in selected_enzymes:
            ki = input[f"ki_{enzyme['id']}"]()
            cmax = input.cmax() / input.mw()
            fup = input.fup()
            fumic = calculate_fumic()
            result = cmax * fup / (ki * fumic)
            if result > 0.02:
                alert = "Risk (R > 0.02)"
            else:
                alert = "" 
            
            note = f"Ki = {ki} μmol/L, fu,mic = {round(fumic,3)}"

            results_ei.append({"Attribute":enzyme["attribute"],"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"],"Risk (R)": round(result, 3),"Alert":alert,"Note":note})
            results_ei_net.append({"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"], "ID": enzyme["id"], "Ki":ki})

        # Convert the results list to a Pandas DataFrame
        results_df_ei = pd.DataFrame(results_ei)
        results_df_ei_net = pd.DataFrame(results_ei_net)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ei["Risk (R)"] = results_df_ei["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 0.02 else "green"};{"font-weight:bold" if x > 0.02 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ei.to_html(escape=False, index=False))


    #Irreversible CYP inhibition - hepatic    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_tdi)
    def result_table_tdi():
        # Collect selected enzymes
        selected_enzymes = [enzyme for enzyme in all_enzymes_tdi if input[f"select_{enzyme['id']}"]()]
        global results_df_tdi 
        global results_df_tdi_net
        results_tdi = []
        results_tdi_net = []

        # Perform calculations for each selected enzyme
        for enzyme in selected_enzymes:
            ki_tdi = input[f"ki_tdi_{enzyme['id']}"]()
            cmax = input.cmax() / input.mw()
            fup = input.fup()
            fumic = calculate_fumic()
            kdeg = input[f"kdeg_{enzyme['id']}"]()
            kinact = input[f"kinact_{enzyme['id']}"]()
            kobs = kinact * 5 * cmax * fup / (ki_tdi * fumic + 5 * cmax * fup)
            result = (kobs + kdeg) / kdeg
            if result > 1.25:
                alert = "Risk (R > 1.25)"
            else:
                alert = ""

            note = f"KI = {ki_tdi} μmol/L, Kinact = {kinact}/h, Kdeg = {kdeg}/h, Kobs = {round(kobs,3)}/h"

            results_tdi.append({"Attribute":enzyme["attribute"],"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"],"Risk (R)": round(result, 3),"Alert":alert,"Note":note})
            results_tdi_net.append({"Enzyme/Transporter": enzyme["name"], "ID": enzyme["id"], "Substrate": enzyme["substrate"], "Kdeg":kdeg,"Kinact":kinact, "KI":ki_tdi})

        # Convert the results list to a Pandas DataFrame
        results_df_tdi = pd.DataFrame(results_tdi)
        results_df_tdi_net = pd.DataFrame(results_tdi_net)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_tdi["Risk (R)"] = results_df_tdi["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 1.25 else "green"};{"font-weight:bold" if x > 1.25 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_tdi.to_html(escape=False, index=False))


    #Competitive CYP inhibition - intestinal    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ei_int)
    def result_table_ei_int():
        # Collect selected enzymes
        selected_enzymes = [enzyme for enzyme in all_enzymes_ei_int if input[f"select_{enzyme['id']}"]()]
        global results_df_ei_int 
        global results_df_ei_int_net
        results_ei_int = []
        results_ei_int_net = []

        # Perform calculations for each selected enzyme
        for enzyme in selected_enzymes:
            ki_int = input[f"ki_{enzyme['id']}"]()
            dose = 1000000 * input.dose() / input.mw()
            fumic = calculate_fumic()
            result = dose / (250 * ki_int * fumic)
            if result > 10:
                alert = "Risk (R > 10)"
            else:
                alert = ""

            note = f"Ki = {ki_int} μmol/L"

            results_ei_int.append({"Attribute":enzyme["attribute"],"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"],"Risk (R)": round(result, 3),"Alert":alert,"Note":note})
            results_ei_int_net.append({"Enzyme/Transporter": enzyme["name"], "ID": enzyme["id"], "Substrate": enzyme["substrate"],"Ki":ki_int})

        # Convert the results list to a Pandas DataFrame
        results_df_ei_int = pd.DataFrame(results_ei_int)
        results_df_ei_int_net = pd.DataFrame(results_ei_int_net)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ei_int["Risk (R)"] = results_df_ei_int["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 10 else "green"};{"font-weight:bold" if x > 10 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ei_int.to_html(escape=False, index=False))


    #Enzyme induction
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ed)
    def result_table_ed():   
        # Collect selected enzymes
        selected_enzymes = [enzyme for enzyme in all_enzymes_ed if input[f"select_{enzyme['id']}"]()]
        global results_df_ed 
        global results_df_ed_net
        results_ed = []
        results_ed_net = []

        # Perform calculations for each selected enzyme
        for enzyme in selected_enzymes:
            emax = input[f"emax_{enzyme['id']}"]()
            ec50 = input[f"ec50_{enzyme['id']}"]()
            cmax = input.cmax() / input.mw()
            fup = input.fup()
            fuhep = calculate_fuhep()
            d = input.d_factor()
            result = 1 / (1 + (10 * d * emax * cmax * fup) / (ec50 * fuhep + 10 * cmax * fup))
            if result < 0.2:
                alert = "Strong inducer (R < 0.2)"
            elif 0.2 <= result < 0.5:
                alert = "Moderate inducer (0.2 ≤ R < 0.5)"
            elif 0.5 <= result < 0.8:
                alert = "Weak inducer (0.5 ≤ R < 0.8)"
            else:
                alert = ""

            note = f"Emax = {emax}, EC50 = {ec50} μmol/L, fu,hep = {round(fuhep,3)}"

            results_ed.append({"Attribute":enzyme["attribute"],"Enzyme/Transporter": enzyme["name"], "Substrate": "-","Risk (R)": round(result, 3),"Alert":alert,"Note":note})
            results_ed_net.append({"Enzyme/Transporter": enzyme["name"], "ID": enzyme["id"],"Emax":emax,"EC50":ec50})

        # Convert the results list to a Pandas DataFrame
        results_df_ed = pd.DataFrame(results_ed)
        results_df_ed_net = pd.DataFrame(results_ed_net)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ed["Risk (R)"] = results_df_ed["Risk (R)"].apply(
            # lambda x: f'<span style="color:black; background-color:{"red" if x < 0.2 else "orange" if 0.2 <= x < 0.5 else "yellow" if 0.5 <= x < 0.8 else "green"}">{x}</span>'
            lambda x: f'<span style="color:{"red" if x < 0.2 else "orange" if 0.2 <= x < 0.5 else "blue" if 0.5 <= x < 0.8 else "green"};{"font-weight:bold" if x < 0.8 else ""}">{x}</span>'
           )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ed.to_html(escape=False, index=False))


    #Net effect
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ne)
    def result_table_ne():   
        # Collect selected enzymes
        selected_enzymes = [enzyme for enzyme in all_enzymes_ne if input[f"select_{enzyme['id']}"]()] + [enzyme for enzyme in all_enzymes_ne_cyp3a4 if input[f"select_{enzyme['id']}"]()] 

        global results_df_ne
        results_ne_ei = []
        results_ne_ed = []
        results_ne_total = []

        # Calculate the sum of fm values
        fm_sum = sum(input[f"fm_{enzyme['id']}"]() for enzyme in selected_enzymes)

        # Check if the sum of fm exceeds 1
        if fm_sum > 1:
            # Raise an alert
            ui.notification_show("The sum of fm values exceeds 1. Please adjust the values.")

        # [I]g, [I]h
        fup = input.fup()
        cmax = input.cmax() / input.mw()
        dose = input.dose() / input.mw()
        fa = input.fa()
        fg = input.fg()
        ka = input.ka()
        qen = input.qen()
        qh = input.qh()
        rb = input.rb()
        fumic = calculate_fumic()
        fuhep = calculate_fuhep()
        i_g = fa * ka * dose * 1000 / qen
        i_h = fup * (cmax + (fa * fg * ka * dose * 1000)/ qh / rb)

        # Iterate over the enzymes and extract the values, setting defaults if DataFrames are not created
        for enzyme in selected_enzymes:   
            d = input.d_factor()
            fm = input[f"fm_{enzyme['id']}"]()
            bg = 1
            cg = 1    

            if "CYP3A" in enzyme["name"]:
                try:
                    emax = results_df_ed_net.loc[results_df_ed_net["Enzyme/Transporter"] == enzyme["name"], "Emax"].values[0]
                except (NameError, KeyError, IndexError):
                    emax = None
                try:
                    ec50 = results_df_ed_net.loc[results_df_ed_net["Enzyme/Transporter"] == enzyme["name"], "EC50"].values[0]
                except (NameError, KeyError, IndexError):
                    ec50 = None

                for substrate in cyp3a_substrates:
                    if input[f'select_{enzyme["id"]}_{substrate["substrate"]}']():
                        try:
                            ki = results_df_ei_net.loc[(results_df_ei_net["Enzyme/Transporter"].str.contains("CYP3A")) & (results_df_ei_net["Substrate"] == substrate["substrate"]), "Ki"].values[0]
                        except (NameError, KeyError, IndexError):
                            ki = None
                        try:
                            kinact = results_df_tdi_net.loc[(results_df_tdi_net["Enzyme/Transporter"].str.contains("CYP3A")) & (results_df_tdi_net["Substrate"] == substrate["substrate"]), "Kinact"].values[0]
                        except (NameError, KeyError, IndexError):
                            kinact = None
                        try:
                            kdeg = results_df_tdi_net.loc[(results_df_tdi_net["Enzyme/Transporter"].str.contains("CYP3A")) & (results_df_tdi_net["Substrate"] == substrate["substrate"]), "Kdeg"].values[0]
                        except (NameError, KeyError, IndexError):
                            kdeg = None
                        try:
                            ki_tdi = results_df_tdi_net.loc[(results_df_tdi_net["Enzyme/Transporter"].str.contains("CYP3A")) & (results_df_tdi_net["Substrate"] == substrate["substrate"]), "KI"].values[0]
                        except (NameError, KeyError, IndexError):
                            ki_tdi = None
                        try:
                            ki_int = results_df_ei_int_net.loc[(results_df_ei_int_net["Enzyme/Transporter"].str.contains("CYP3A")) & (results_df_ei_int_net["Substrate"] == substrate["substrate"]), "Ki"].values[0]
                        except (NameError, KeyError, IndexError):
                            ki_int = None

                        ag = 1 if ki_int is None else 1 / (1 + (i_g / (fumic * ki_int))) 
                        ah = 1 if ki is None else 1 / (1 + (i_h / (fumic * ki)))
                        bh = 1 if kinact is None or kdeg is None or ki_tdi is None else kdeg / (kdeg + (i_h * kinact / (i_h + fumic * ki_tdi)))
                        ch = 1 if emax is None or ec50 is None else 1 + (d * emax * i_h) / (i_h + fuhep * ec50)

                        ag_ed = 1
                        ah_ed = 1
                        bh_ed = 1
                        ch_ei = 1

                        result_ei = (1/(ag * bg * cg * (1-fg) + fg)) * (1 / (ah * bh * ch_ei * fm + (1 - fm)))                        
                        result_ed = (1/(ag_ed * bg * cg * (1-fg) + fg)) * (1 / (ah_ed * bh_ed * ch * fm + (1 - fm)))                        
                        result_total = (1/(ag * bg * cg * (1-fg) + fg)) * (1 / (ah * bh * ch * fm + (1 - fm)))
                        
                        if result_ei > 1.25:
                            alert_ei = "Risk (R > 1.25)"
                        else:
                            alert_ei = ""
                        
                        if result_ed < 0.8:
                            alert_ed = "Risk (R < 0.8)"
                        else:
                            alert_ed = ""

                        if result_total < 0.8:
                            alert_total = "Risk (R < 0.8)"
                        elif result_total > 1.25:
                            alert_total = "Risk (R > 1.25)"
                        else:
                            alert_total = ""

                        i_g_r3 = round(i_g,3)    
                        i_h_r3 = round(i_h,3) 
                        ag_r3 = round(ag,3) 
                        bg_r3 = round(bg,3) 
                        bg_r3 = round(bg,3) 
                        cg_r3 = round(cg,3) 
                        ah_r3 = round(ah,3) 
                        bh_r3 = round(bh,3) 
                        ch_r3 = round(ch,3) 

                        note_ei = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_r3}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_r3}, Bh = {bh_r3} and Ch = {ch_ei}"
                        note_ed = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_ed}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_ed}, Bh = {bh_ed} and Ch = {ch_r3}"
                        note_total = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_r3}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_r3}, Bh = {bh_r3} and Ch = {ch_r3}"
                
                        results_ne_ei.append({"Attribute":f'{enzyme["attribute"]} (inhibition only)',"Enzyme/Transporter": enzyme["name"], "Substrate": substrate["substrate"], "Risk (R)": round(result_ei, 3),"Alert":alert_ei,"Note":note_ei})
                        results_ne_ed.append({"Attribute":f'{enzyme["attribute"]} (induction only)',"Enzyme/Transporter": enzyme["name"], "Substrate": substrate["substrate"], "Risk (R)": round(result_ed, 3),"Alert":alert_ed,"Note":note_ed})
                        results_ne_total.append({"Attribute":f'{enzyme["attribute"]} (AUCR)',"Enzyme/Transporter": enzyme["name"], "Substrate": substrate["substrate"], "Risk (R)": round(result_total, 3),"Alert":alert_total,"Note":note_total})

            else:
                ki_int = None # Other than CYP3A no ki_int available
                try:
                    ki = results_df_ei_net.loc[results_df_ei_net["Enzyme/Transporter"] == enzyme["name"], "Ki"].values[0]
                except (NameError, KeyError, IndexError):
                    ki = None
                try:
                    kinact = results_df_tdi_net.loc[results_df_tdi_net["Enzyme/Transporter"] == enzyme["name"], "Kinact"].values[0]
                except (NameError, KeyError, IndexError):
                    kinact = None
                try:
                    kdeg = results_df_tdi_net.loc[results_df_tdi_net["Enzyme/Transporter"] == enzyme["name"], "Kdeg"].values[0]
                except (NameError, KeyError, IndexError):
                    kdeg = None
                try:
                    ki_tdi = results_df_tdi_net.loc[results_df_tdi_net["Enzyme/Transporter"] == enzyme["name"], "KI"].values[0]
                except (NameError, KeyError, IndexError):
                    ki_tdi = None
                try:
                    emax = results_df_ed_net.loc[results_df_ed_net["Enzyme/Transporter"] == enzyme["name"], "Emax"].values[0]
                except (NameError, KeyError, IndexError):
                    emax = None
                try:
                    ec50 = results_df_ed_net.loc[results_df_ed_net["Enzyme/Transporter"] == enzyme["name"], "EC50"].values[0]
                except (NameError, KeyError, IndexError):
                    ec50 = None
                
                ag = 1 if ki_int is None else 1 / (1 + (i_g / (fumic * ki_int))) 
                ah = 1 if ki is None else 1 / (1 + (i_h / (fumic * ki)))
                bh = 1 if kinact is None or kdeg is None or ki_tdi is None else kdeg / (kdeg + (i_h * kinact / (i_h + fumic * ki_tdi)))
                ch = 1 if emax is None or ec50 is None else 1 + (d * emax * i_h) / (i_h + fuhep * ec50)

                ag_ed = 1
                ah_ed = 1
                bh_ed = 1
                ch_ei = 1

                result_ei = (1/(ag * bg * cg * (1-fg) + fg)) * (1 / (ah * bh * ch_ei * fm + (1 - fm)))                        
                result_ed = (1/(ag_ed * bg * cg * (1-fg) + fg)) * (1 / (ah_ed * bh_ed * ch * fm + (1 - fm)))                        
                result_total = (1/(ag * bg * cg * (1-fg) + fg)) * (1 / (ah * bh * ch * fm + (1 - fm)))
                
                if result_ei > 1.25:
                    alert_ei = "Risk (R > 1.25)"
                else:
                    alert_ei = ""
                
                if result_ed < 0.8:
                    alert_ed = "Risk (R < 0.8)"
                else:
                    alert_ed = ""

                if result_total < 0.8:
                    alert_total = "Risk (R < 0.8)"
                elif result_total > 1.25:
                    alert_total = "Risk (R > 1.25)"
                else:
                    alert_total = ""

                i_g_r3 = round(i_g,3)    
                i_h_r3 = round(i_h,3) 
                ag_r3 = round(ag,3) 
                bg_r3 = round(bg,3) 
                bg_r3 = round(bg,3) 
                cg_r3 = round(cg,3) 
                ah_r3 = round(ah,3) 
                bh_r3 = round(bh,3) 
                ch_r3 = round(ch,3) 

                note_ei = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_r3}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_r3}, Bh = {bh_r3} and Ch = {ch_ei}"
                note_ed = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_ed}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_ed}, Bh = {bh_ed} and Ch = {ch_r3}"
                note_total = f"[I]g = {i_g_r3} μmol/L, Ag = {ag_r3}, Bg = {bg_r3}, Cg = {cg_r3}, [I]h = {i_h_r3} μmol/L, fm = {fm}, Ah = {ah_r3}, Bh = {bh_r3} and Ch = {ch_r3}"

                results_ne_ei.append({"Attribute":f'{enzyme["attribute"]} (inhibition only)',"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"], "Risk (R)": round(result_ei, 3),"Alert":alert_ei,"Note":note_ei})
                results_ne_ed.append({"Attribute":f'{enzyme["attribute"]} (induction only)',"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"], "Risk (R)": round(result_ed, 3),"Alert":alert_ed,"Note":note_ed})
                results_ne_total.append({"Attribute":f'{enzyme["attribute"]} (AUCR)',"Enzyme/Transporter": enzyme["name"], "Substrate": enzyme["substrate"], "Risk (R)": round(result_total, 3),"Alert":alert_total,"Note":note_total})

        # Convert the results list to a Pandas DataFrame
        combined_results = results_ne_ei + results_ne_ed + results_ne_total
        results_df_ne_beforesort = pd.DataFrame(combined_results)
        results_df_ne = results_df_ne_beforesort.sort_values(by=["Enzyme/Transporter", "Attribute"])

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ne["Risk (R)"] = results_df_ne["Risk (R)"].apply(
            # lambda x: f'<span style="color:black; background-color:{"red" if x < 0.2 else "orange" if 0.2 <= x < 0.5 else "yellow" if 0.5 <= x < 0.8 else "green"}">{x}</span>'
            lambda x: f'<span style="color:{"red" if x < 0.8 else "red" if x > 1.25 else "green"};{"font-weight:bold" if x < 0.8 or x > 1.25 else ""}">{x}</span>'
           )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ne.to_html(escape=False, index=False))



    #Transporter inhibition - intestinal efflux    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ti_ie)
    def result_table_ti_ie():
        # Collect selected enzymes
        selected_transporters = [transporter for transporter in all_transporters_ie if input[f"select_{transporter['id']}"]()]
        global results_df_ti_ie
        results_ti_ie = []

        # Perform calculations for each selected enzyme
        for transporter in selected_transporters:
            ic50 = input[f"ic50_{transporter['id']}"]()
            dose = 1000000 * input.dose() / input.mw()
            fuinc = input.fuinc()
            result = dose / (250 * ic50 * fuinc)
            if result > 10:
                alert = "Risk (R > 10)"
            else:
                alert = ""

            note = f"IC50 or Ki = {ic50} μmol/L"

            results_ti_ie.append({"Attribute":transporter["attribute"],"Enzyme/Transporter": transporter["name"], "Substrate": "-", "Risk (R)": round(result, 3),"Alert":alert,"Note":note})

        # Convert the results list to a Pandas DataFrame
        results_df_ti_ie = pd.DataFrame(results_ti_ie)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ti_ie["Risk (R)"] = results_df_ti_ie["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 10 else "green"};{"font-weight:bold" if x > 10 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ti_ie.to_html(escape=False, index=False))


    #Transporter inhibition - hepatic uptake    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ti_hu)
    def result_table_ti_hu():
        # Collect selected enzymes
        selected_transporters = [transporter for transporter in all_transporters_hu if input[f"select_{transporter['id']}"]()]
        global results_df_ti_hu
        results_ti_hu = []

        # Perform calculations for each selected enzyme
        for transporter in selected_transporters:
            ic50 = input[f"ic50_{transporter['id']}"]()
            fup = input.fup()
            fuinc = input.fuinc()
            cmax = input.cmax() / input.mw()
            dose = input.dose() / input.mw()
            fa = input.fa()
            fg = input.fg()
            ka = input.ka()
            qh = input.qh()
            rb = input.rb()
            result = (cmax + (fa * fg * ka * dose * 1000 / qh / rb)) * fup / (ic50 * fuinc)
            if result > 0.1:
                alert = "Risk (R > 0.1)"
            else:
                alert = ""

            note = f"IC50 or Ki = {ic50} μmol/L"

            results_ti_hu.append({"Attribute":transporter["attribute"],"Enzyme/Transporter": transporter["name"], "Substrate": "-", "Risk (R)": round(result, 3),"Alert":alert,"Note":note})

        # Convert the results list to a Pandas DataFrame
        results_df_ti_hu = pd.DataFrame(results_ti_hu)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ti_hu["Risk (R)"] = results_df_ti_hu["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 0.1 else "green"};{"font-weight:bold" if x > 0.1 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ti_hu.to_html(escape=False, index=False))


    #Transporter inhibition - renal uptake    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ti_ru)
    def result_table_ti_ru():
        # Collect selected enzymes
        selected_transporters = [transporter for transporter in all_transporters_ru if input[f"select_{transporter['id']}"]()]
        global results_df_ti_ru
        results_ti_ru = []

        # Perform calculations for each selected enzyme
        for transporter in selected_transporters:
            ic50 = input[f"ic50_{transporter['id']}"]()
            fup = input.fup()
            fuinc = input.fuinc()
            cmax = input.cmax() / input.mw()
            result = cmax * fup / (ic50 * fuinc)
            if result > 0.1:
                alert = "Risk (R > 0.1)"
            else:
                alert = ""

            note = f"IC50 or Ki = {ic50} μmol/L"

            results_ti_ru.append({"Attribute":transporter["attribute"],"Enzyme/Transporter": transporter["name"], "Substrate": "-", "Risk (R)": round(result, 3),"Alert":alert,"Note":note})

        # Convert the results list to a Pandas DataFrame
        results_df_ti_ru = pd.DataFrame(results_ti_ru)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ti_ru["Risk (R)"] = results_df_ti_ru["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 0.1 else "green"};{"font-weight:bold" if x > 0.1 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ti_ru.to_html(escape=False, index=False))


    #Transporter inhibition - renal efflux    
    # @output
    @render.ui
    @reactive.event(input.calculate_all_ti_re)
    def result_table_ti_re():
        # Collect selected enzymes
        selected_transporters = [transporter for transporter in all_transporters_re if input[f"select_{transporter['id']}"]()]
        global results_df_ti_re
        results_ti_re = []

        # Perform calculations for each selected enzyme
        for transporter in selected_transporters:
            ic50 = input[f"ic50_{transporter['id']}"]()
            fup = input.fup()
            fuinc = input.fuinc()
            cmax = input.cmax() / input.mw()
            result = cmax * fup / (ic50 * fuinc)
            if result > 0.02:
                alert = "Risk (R > 0.02)"
            else:
                alert = ""

            note = f"IC50 or Ki = {ic50} μmol/L"

            results_ti_re.append({"Attribute":transporter["attribute"],"Enzyme/Transporter": transporter["name"], "Substrate": "-", "Risk (R)": round(result, 3),"Alert":alert,"Note":note})

        # Convert the results list to a Pandas DataFrame
        results_df_ti_re = pd.DataFrame(results_ti_re)

        # Apply conditional formatting to only the "Risk (R)" column
        results_df_ti_re["Risk (R)"] = results_df_ti_re["Risk (R)"].apply(
            lambda x: f'<span style="color:{"red" if x > 0.02 else "green"};{"font-weight:bold" if x > 0.02 else ""}">{x}</span>'
        )

        # Convert the DataFrame to HTML
        return ui.HTML(results_df_ti_re.to_html(escape=False, index=False))


    # Summary
    # @output
    @render.ui
    @reactive.event(input.summary)
    def summary_table(): 
        global unified_df1, unified_df2, unified_df3, unified_df4, user_memo_df
        # List of DataFrame names for each table
        df_names1 = ["results_df_ei", "results_df_tdi", "results_df_ei_int"]
        df_names2 = ["results_df_ed"]
        df_names3 = ["results_df_ne"]
        df_names4 = ["results_df_ti_ie", "results_df_ti_hu", "results_df_ti_ru", "results_df_ti_re"]
        
        # Function to concatenate DataFrames from a list of names
        def concatenate_dfs(df_names):
            dfs = []
            for name in df_names:
                try:
                    df = eval(name)
                    if df is not None and not df.empty:
                        dfs.append(df)
                except NameError:
                    pass
            return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

        # Create three separate DataFrames
        unified_df1 = concatenate_dfs(df_names1)
        unified_df2 = concatenate_dfs(df_names2)
        unified_df3 = concatenate_dfs(df_names3)
        unified_df4 = concatenate_dfs(df_names4)

        # Convert the DataFrames to HTML
        html_table1 = unified_df1.to_html(escape=False, index=False)
        html_table2 = unified_df2.to_html(escape=False, index=False)
        html_table3 = unified_df3.to_html(escape=False, index=False)
        html_table4 = unified_df4.to_html(escape=False, index=False)

        # Combine the HTML tables with separators
        combined_html = f"<h3>Enzyme inhibition</h3>{html_table1}<br><br><h3>Enzyme induction</h3>{html_table2}<br><br><h3>Net effect</h3>{html_table3}<br><br><h3>Transporter inhibition</h3>{html_table4}"

        # Convert user_memo to DataFrame
        user_memo = input.user_memo()
        user_memo_lines = user_memo.split('\n')  # Split the text into lines
        user_memo_data = [line.split(',') for line in user_memo_lines]  # Split each line into columns
        user_memo_df = pd.DataFrame(user_memo_data[1:], columns=user_memo_data[0])  # Create DataFrame

        return ui.HTML(combined_html)

    # Output excel file
    # Function to strip HTML tags and extract styles
    def extract_text_and_style(html):
        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text()
        style = {}
        if soup.span:
            style['bold'] = 'font-weight:bold' in soup.span.get('style', '')
            style['color'] = soup.span.get('style', '').split('color:')[-1].split(';')[0] if 'color:' in soup.span.get('style', '') else None
            style['bgcolor'] = soup.span.get('style', '').split('background-color:')[-1].split(';')[0] if 'background-color:' in soup.span.get('style', '') else None
        return text, style

    # Function to convert color to aRGB hex format
    def to_argb_hex(color):
        if color.startswith('#'):
            color = color[1:]
        if len(color) == 6:
            color = 'FF' + color  # Add alpha value
        return color.upper()

    # Mapping of named colors to hex values
    named_colors = {
        'red': 'FF0000',
        'green': '00FF00',
        'blue': '0000FF',
        'black': '000000',
        'white': 'FFFFFF',
        'yellow': 'FFFF00',
        'cyan': '00FFFF',
        'magenta': 'FF00FF',
        'orange': 'FFA500'
    }

    # Function to validate and convert color
    def validate_and_convert_color(color):
        if color:
            if color in named_colors:
                color = named_colors[color]
            elif color.startswith('rgb'):
                # Convert rgb(r, g, b) to hex
                color = color.replace('rgb(', '').replace(')', '')
                r, g, b = map(int, color.split(','))
                color = f'FF{r:02X}{g:02X}{b:02X}'
            elif color.startswith('#'):
                color = to_argb_hex(color)
            return color
        return None

    @output
    @render.download(filename=f"statDDI_{datetime.now().strftime('%Y%m%d')}.xlsx")
    def handle_download_xlsx():
        # Generate the Excel content
        def excel_content():
            excel_buffer = io.BytesIO()
            
            # Create a new workbook and select the active worksheet
            workbook = Workbook()
            worksheet = workbook.active
            
            # Ensure input parameters are correctly referenced
            drug_product = input.cmp()  
            mw = input.mw()
            cmax = input.cmax()
            fup = input.fup()
            dose = input.dose() 
            ka = input.ka()
            fa = input.fa()
            fg = input.fg()
            qh = input.qh()
            rb = input.rb()
            d_factor = input.d_factor()
            qen = input.qen()
            fuinc = input.fuinc()

            # Write the profile information to the first row
            profile_text = f'Compound: {drug_product}'
            input_values = f'MW = {mw} g/mol, Cmax = {cmax} ng/mL, fu,p = {fup}, Dose = {dose} mg, ka = {ka}/h, Fa = {fa}, Fg = {fg}, Qh = {qh} L/h, Qen = {qen} L/h, Rb = {rb}, d factor = {d_factor}, fu,inc = {fuinc}'

            cell = worksheet.cell(row=1, column=1, value=profile_text)
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(wrap_text=True)

            cell = worksheet.cell(row=2, column=1, value=input_values)
            cell.alignment = Alignment(wrap_text=False)

            # Function to write a DataFrame to the worksheet with a title
            def write_df_with_title(df, title, start_row):
                worksheet.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12)
                for col_num, column_title in enumerate(df.columns, 1):
                    cell = worksheet.cell(row=start_row + 1, column=col_num, value=column_title)
                    cell.font = Font(bold=True)
                for r_idx, row in enumerate(df.itertuples(index=False), start_row + 2):
                    for c_idx, value in enumerate(row, 1):
                        if isinstance(value, str) and '<span' in value:
                            text, style = extract_text_and_style(value)
                            cell = worksheet.cell(row=r_idx, column=c_idx, value=text)
                            font_args = {}
                            if style.get('bold'):
                                font_args['bold'] = True
                            if style.get('color'):
                                color = validate_and_convert_color(style['color'])
                                if color:
                                    font_args['color'] = Color(rgb=color)
                            cell.font = Font(**font_args)
                        else:
                            worksheet.cell(row=r_idx, column=c_idx, value=value)
                return start_row + len(df) + 3  # Return the next starting row

            # Write each DataFrame with its title
            next_row = 4
            next_row = write_df_with_title(unified_df1, "Enzyme inhibition", next_row)
            next_row = write_df_with_title(unified_df2, "Enzyme induction", next_row)
            next_row = write_df_with_title(unified_df3, "Net effect", next_row)
            next_row = write_df_with_title(unified_df4, "Transporter inhibition", next_row)
            write_df_with_title(user_memo_df, "User memo", next_row)
            
            # Adjust column widths
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
            
            worksheet.column_dimensions['A'].width = 50

            # Save the workbook to the buffer
            workbook.save(excel_buffer)
            excel_buffer.seek(0)  # Rewind the buffer for reading
            return excel_buffer.getvalue()

        try:
            # Yield the Excel content for downloading
            yield excel_content()
        except asyncio.CancelledError:
            print("Download was cancelled.")


    class PDF(FPDF, HTMLMixin):
        def header(self):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, 'Summary report (Static DDI Risk Calculator)', 0, 1, 'C')

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    @output
    @render.download(filename=f"statDDI_{datetime.now().strftime('%Y%m%d')}.pdf")
    def handle_download_pdf():
        def pdf_content():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=10 * mm, rightMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
            elements = []

            # Styles
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            normal_style = styles['Normal']
            bold_style = ParagraphStyle(name='Bold', parent=styles['Normal'], fontName='Helvetica-Bold')

            # Header
            header = Paragraph('Summary report (Static DDI Risk Calculator)', title_style)
            elements.append(header)
            elements.append(Spacer(1, 0.5 * cm))

            # Input parameters
            drug_product = input.cmp()
            mw = input.mw()
            cmax = input.cmax()
            fup = input.fup()
            dose = input.dose()
            ka = input.ka()
            fa = input.fa()
            fg = input.fg()
            qh = input.qh()
            rb = input.rb()
            d_factor = input.d_factor()
            qen = input.qen()
            fuinc = input.fuinc()

            # Profile information
            profile_text = f'Compound: {drug_product}'
            input_values = f'MW = {mw} g/mol, Cmax = {cmax} ng/mL, fu,p = {fup}, Dose = {dose} mg, ka = {ka}/h, Fa = {fa}, Fg = {fg}, Qh = {qh} L/h, Qen = {qen} L/h, Rb = {rb}, d factor = {d_factor}, fu,inc = {fuinc}'
            elements.append(Paragraph(profile_text, bold_style))
            elements.append(Paragraph(input_values, normal_style))
            elements.append(Spacer(1, 0.5 * cm))

            # Function to write DataFrame to PDF
            def write_df_to_pdf(df, title, add_title=True, add_background=True, add_grid=True):
                if add_title:
                    elements.append(Paragraph(title, bold_style))
                if df.empty:
                    elements.append(Paragraph("No data available", normal_style))
                    elements.append(Spacer(1, 1 * cm))
                    return
                data = [df.columns.tolist()] + df.values.tolist()
                formatted_data = []
                for row in data:
                    formatted_row = []
                    for cell in row:
                        cell_html = str(cell)
                        soup = BeautifulSoup(cell_html, 'html.parser')
                        text_only = soup.get_text() #remove HTML tag
                        formatted_row.append(Paragraph(text_only, normal_style))
                        # for span in soup.find_all('span'):
                        #     style = span.get('style', '')
                        #     if 'color' in style:
                        #         color = style.split('color:')[1].split(';')[0].strip()
                        #         span.insert_before(f'<font color="{color}">')
                        #         span.insert_after('</font>')
                        #     if 'font-weight: bold' in style:
                        #         span.insert_before('<b>')
                        #         span.insert_after('</b>')
                        #     span.unwrap()
                        # formatted_row.append(Paragraph(str(soup), normal_style))
                    formatted_data.append(formatted_row)
                table = Table(formatted_data, repeatRows=1)
                table_style = [
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ]
                if add_background:
                    table_style.extend([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black), 
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ])
                if add_grid:
                    table_style.append(('GRID', (0, 0), (-1, -1), 1, colors.black))
                table.setStyle(TableStyle(table_style))
                elements.append(table)
                elements.append(Spacer(1, 0.5 * cm))

            # Write DataFrames to PDF
            write_df_to_pdf(unified_df1, "Enzyme inhibition")
            write_df_to_pdf(unified_df2, "Enzyme induction")
            write_df_to_pdf(unified_df3, "Net effect")
            write_df_to_pdf(unified_df4, "Transporter inhibition")
            write_df_to_pdf(user_memo_df, "User memo", add_title=True, add_background=False, add_grid=False)

            # Build PDF
            doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
            buffer.seek(0)
            return buffer.getvalue()

        # Initialize page count
        page_count = [0]

        def add_page_number(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            page_count[0] += 1
            page_number_text = f"Page {page_count[0]}"
            page_width = doc.pagesize[0]  
            canvas.drawCentredString(page_width / 2.0, 10 * mm, page_number_text)
            canvas.restoreState()

        try:
            yield pdf_content()
        except asyncio.CancelledError:
            print("Download was cancelled.")

    # Glossary
    @output
    @render.table
    def glossary_table():
        return glossary_df

    results = reactive.Value({
        "enzyme_inhibition": None,
        "enzyme_induction": None,
        "net_effect": None,
        "transporter_inhibition": None,

    })


# Create the app
app = App(app_ui, server)
