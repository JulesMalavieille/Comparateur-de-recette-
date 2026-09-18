"""
Created on Tue Sep 15 18:42:45 2026

@author: Jules Malavieille
"""
    
    
import streamlit as st
import pandas as pd 
import re
import io 
from datetime import datetime


FUNCTIONAL_CLASSES = {
    "acidifiant": "acidifiant",
    "acidifiants": "acidifiant",

    "humectant": "humectant",
    "humectants": "humectant",

    "émulsifiant": "émulsifiant",
    "émulsifiants": "émulsifiant",

    "colorant": "colorant",
    "colorants": "colorant",

    "gélifiant": "gélifiant",
    "gélifiants": "gélifiant",

    "conservateur": "conservateur",
    "conservateurs": "conservateur",

    "antioxydant": "antioxydant",
    "antioxydants": "antioxydant",

    "édulcorant": "édulcorant",
    "édulcorants": "édulcorant",

    "correcteur d'acidité": "correcteur d'acidité",
    "correcteurs d'acidité": "correcteur d'acidité",

    "agent d'enrobage": "agent d'enrobage",
    "agents d'enrobage": "agent d'enrobage"}


# =========================================================
# CONFIGURATION DE LA PAGE
# =========================================================

st.set_page_config(
    page_title="Comparateur de recettes",
    page_icon="🔎",
    layout="wide")

st.markdown("""<style>.block-container {max-width: 1200px; padding-top: 2rem; padding-bottom: 4rem;}</style>""", unsafe_allow_html=True)


# =========================================================
# FONCTIONS
# =========================================================

def clean_text(text):

    text = re.sub(
        r"^\s*ingr[eé]dients?\s*:\s*",
        "",
        text,
        flags=re.IGNORECASE)

    return text.strip()



def detect_separator(text):

    depth = 0

    for char in text:

        if char == "(":
            depth += 1

        elif char == ")":
            depth -= 1

        elif char == ";" and depth == 0:
            return ";"

    return ","



def comparison_key(ingredient):

    if ":" not in ingredient:
        return ingredient

    category, content = ingredient.split(":", 1)

    category_clean = category.strip().lower()

    if category_clean in FUNCTIONAL_CLASSES:
        category_clean = FUNCTIONAL_CLASSES[category_clean]

    return f"{category_clean} : {content.strip()}"



def split_ingredients(text, separator):

    ingredients = []
    element = ""
    depth = 0

    for char in text:

        if char == "(":
            depth += 1

        elif char == ")":
            depth -= 1

        if char == separator and depth == 0:

            if element.strip():
                ingredients.append(element.strip())

            element = ""

        else:
            element += char

    if element.strip():
        ingredients.append(element.strip())

    return ingredients



def prepare_ingredients(text):

    text = clean_text(text)

    separator = detect_separator(text)

    blocks = split_ingredients(text, separator)

    ingredients = []

    for block in blocks:

        # Cas d'une catégorie fonctionnelle
        if ":" in block:

            category, content = block.split(":", 1)

            category_clean = category.strip().lower()

            if category_clean in FUNCTIONAL_CLASSES:

                category_key = FUNCTIONAL_CLASSES[
                    category_clean]

                # Si la recette utilise ;
                # les éléments de la catégorie sont
                # généralement séparés par des virgules
                subingredients = split_ingredients(
                    content,
                    ",")

                for subingredient in subingredients:

                    subingredient = subingredient.strip()

                    ingredients.append({
                        "key": (
                            "functional",
                            category_key,
                            subingredient
                        ),
                        "display": subingredient
                    })

                continue

        # Ingrédient classique
        ingredients.append({
            "key": (
                "ingredient",
                block
            ),
            "display": block})

    return ingredients



def compare_recipes(textA, textB):

    A = prepare_ingredients(textA)
    B = prepare_ingredients(textB)

    keys_A = [ingredient["key"] for ingredient in A]
    keys_B = [ingredient["key"] for ingredient in B]

    removed = [
        ingredient["display"]
        for ingredient in A
        if ingredient["key"] not in keys_B]

    added = [
        ingredient["display"]
        for ingredient in B
        if ingredient["key"] not in keys_A]

    unchanged = [
        ingredient["display"]
        for ingredient in A
        if ingredient["key"] in keys_B]

    return removed, added, unchanged



def clear_recipes():
    st.session_state["recipe_a"] = ""
    st.session_state["recipe_b"] = ""



def create_excel(nameA, nameB, textA, textB, removed, added, unchanged):

    # -----------------------------------------------------
    # Mise à la même longueur des 3 listes
    # -----------------------------------------------------

    max_len = max(len(removed), len(added), len(unchanged), 1)

    removed_col = removed + [""] * (max_len - len(removed))
    added_col = added + [""] * (max_len - len(added))
    unchanged_col = unchanged + [""] * (max_len - len(unchanged))

    df_results = pd.DataFrame({
        "Retirés": removed_col,
        "Ajoutés": added_col,
        "Inchangés": unchanged_col
    })

    # -----------------------------------------------------
    # Création du fichier Excel en mémoire
    # -----------------------------------------------------

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        workbook = writer.book

        # =================================================
        # FORMATS
        # =================================================

        title_format = workbook.add_format({
            "bold": True,
            "font_size": 18,
            "align": "left",
            "valign": "vcenter",
            "font_color": "#1E2925"
        })

        subtitle_format = workbook.add_format({
            "font_size": 10,
            "font_color": "#6B7470"
        })

        label_format = workbook.add_format({
            "bold": True,
            "font_color": "#1E2925"
        })

        value_format = workbook.add_format({
            "font_color": "#1E2925"
        })

        removed_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F4CCCC",
            "font_color": "#8A1C1C",
            "border": 1
        })

        added_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9EAD3",
            "font_color": "#1F6F43",
            "border": 1
        })

        unchanged_header = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#E7E7E7",
            "font_color": "#555555",
            "border": 1
        })

        removed_cell = workbook.add_format({
            "bg_color": "#FCE8E8",
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })

        added_cell = workbook.add_format({
            "bg_color": "#EAF4E7",
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })

        unchanged_cell = workbook.add_format({
            "bg_color": "#F3F3F3",
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })

        source_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9E2DD",
            "border": 1,
            "align": "center"
        })

        source_cell = workbook.add_format({
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })


        # =================================================
        # FEUILLE 1 : COMPARAISON
        # =================================================

        worksheet = workbook.add_worksheet("Comparaison")

        # Titre
        worksheet.merge_range(
            "A1:C1",
            "Comparaison de recettes",
            title_format
        )

        worksheet.write(
            "A2",
            f"Rapport généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
            subtitle_format
        )

        # Informations sur les recettes
        worksheet.write("A4", "Recette initiale :", label_format)
        worksheet.write("B4", nameA, value_format)

        worksheet.write("A5", "Nouvelle recette :", label_format)
        worksheet.write("B5", nameB, value_format)

        # Résumé chiffré
        worksheet.write("A7", f"Retirés : {len(removed)}", removed_header)
        worksheet.write("B7", f"Ajoutés : {len(added)}", added_header)
        worksheet.write("C7", f"Inchangés : {len(unchanged)}", unchanged_header)

        # Titres des colonnes
        worksheet.write("A9", "RETIRÉS", removed_header)
        worksheet.write("B9", "AJOUTÉS", added_header)
        worksheet.write("C9", "INCHANGÉS", unchanged_header)

        # Données
        for row, ingredient in enumerate(removed_col, start=9):
            worksheet.write(row, 0, ingredient, removed_cell)

        for row, ingredient in enumerate(added_col, start=9):
            worksheet.write(row, 1, ingredient, added_cell)

        for row, ingredient in enumerate(unchanged_col, start=9):
            worksheet.write(row, 2, ingredient, unchanged_cell)

        # Largeurs
        worksheet.set_column("A:C", 45)

        # Hauteur de la ligne du titre
        worksheet.set_row(0, 28)

        # Figer le haut de la liste
        worksheet.freeze_panes(9, 0)

        # Masquer le quadrillage Excel
        worksheet.hide_gridlines(2)


        # =================================================
        # FEUILLE 2 : RECETTES SOURCES
        # =================================================

        source_sheet = workbook.add_worksheet("Recettes sources")

        source_sheet.merge_range(
            "A1:B1",
            "Recettes originales",
            title_format
        )

        source_sheet.write("A3", nameA, source_header)
        source_sheet.write("B3", nameB, source_header)

        source_sheet.write("A4", textA, source_cell)
        source_sheet.write("B4", textB, source_cell)

        source_sheet.set_column("A:B", 70)

        # Grande hauteur pour lire les recettes
        source_sheet.set_row(3, 120)

        source_sheet.hide_gridlines(2)


    output.seek(0)

    return output




# =========================================================
# EN-TÊTE
# =========================================================

st.markdown("# Comparateur de recettes")

st.markdown(
    """ Comparez deux listes d'ingrédients et identifiez immédiatement
    les éléments **ajoutés**, **retirés** et **inchangés**.""")

st.divider()


# =========================================================
# BOUTON EXEMPLE
# =========================================================


if st.button("Charger un exemple"):
    st.session_state["recipe_a"] = (
        "tomates 83%, fromage (lait, sel, amidon), "
        "sel, correcteur d'acidité : acide citrique, basilic")

    st.session_state["recipe_b"] = (
        "tomates 61%, fromage (lait, sel, vinaigre), "
        "sucre, correcteur d'acidité : acide citrique, basilic")
    
    
# =========================================================
# BOUTON VIDER
# =========================================================
st.button(
    "Vider les recettes",
    on_click=clear_recipes,
    icon=":material/delete:")
    


# =========================================================
# ZONES DE TEXTE
# =========================================================
nameA = st.text_input(
    "Nom de la recette A",
    value="Recette A")

nameB = st.text_input(
    "Nom de la recette B",
    value="Recette B")


colA, colB = st.columns(2, gap="large")


with colA:

    with st.container(border=True):

        st.markdown(f"### {nameA}")
        st.caption("Version initiale")

        textA = st.text_area(
            nameA,
            height=250,
            placeholder="Collez la première liste d'ingrédients...",
            label_visibility="collapsed", 
            key="recipe_a")


with colB:

    with st.container(border=True):

        st.markdown(f"### {nameB}")
        st.caption("Nouvelle version")

        textB = st.text_area(
            nameB,
            height=250,
            placeholder="Collez la deuxième liste d'ingrédients...",
            label_visibility="collapsed", 
            key="recipe_b")


# =========================================================
# BOUTON COMPARE
# =========================================================

st.write("")

left, center, right = st.columns([1.5, 2, 1.5])

with center:

    compare = st.button("Comparer les recettes", type="primary", width="stretch", icon=":material/compare_arrows:")
    



# =========================================================
# RÉSULTATS
# =========================================================

if compare:

    if textA.strip() == "" or textB.strip() == "":

        st.warning("Veuillez entrer deux listes d'ingrédients.")

    else:

        removed, added, unchanged = compare_recipes(textA, textB)

        excel_file = create_excel(
            nameA,
            nameB,
            textA,
            textB,
            removed,
            added,
            unchanged)

        # ---------------------------------------------
        # DESCENDRE AUTOMATIQUEMENT AUX RÉSULTATS
        # ---------------------------------------------

        st.html(
            """
            <div id="resultats"></div>

            <script>
                setTimeout(() => {
                    const element = document.getElementById("resultats");

                    if (element) {
                        element.scrollIntoView({
                            behavior: "smooth",
                            block: "start"
                        });
                    }
                }, 200);
            </script>
            """,
            unsafe_allow_javascript=True)

        st.divider()

        st.subheader("Résultat de la comparaison")


        # ---------------------------------------------
        # RÉSUMÉ
        # ---------------------------------------------

        metric1, metric2, metric3 = st.columns(3)

        metric1.metric("Retirés", len(removed))
        metric2.metric("Ajoutés", len(added))
        metric3.metric("Inchangés", len(unchanged))

        st.write("")


        # ---------------------------------------------
        # DÉTAILS
        # ---------------------------------------------

        col_removed, col_added = st.columns(2, gap="large")


        with col_removed:

            with st.container(border=True):

                st.subheader("Éléments retirés")

                if removed:

                    for ingredient in removed:
                        st.error(f"− {ingredient}")

                else:
                    st.write(
                        "Aucun ingrédient n'a été retiré de la recette.")


        with col_added:

            with st.container(border=True):

                st.subheader("Éléments ajoutés")

                if added:

                    for ingredient in added:
                        st.success(f"+ {ingredient}")

                else:
                    st.write(
                        "Aucun ingrédient n'a été ajouté à la recette.")


        # ---------------------------------------------
        # AUCUN CHANGEMENT
        # ---------------------------------------------

        if len(removed) == 0 and len(added) == 0:

            st.success(
                "Les deux listes d'ingrédients sont identiques.")


        # ---------------------------------------------
        # RAPPORT EXCEL
        # ---------------------------------------------

        st.write("")

        left, center, right = st.columns([1.5, 2, 1.5])

        with center:

            st.download_button(
                label="Télécharger les résultats (.xlsx)",
                data=excel_file,
                file_name="comparaison_recettes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                icon=":material/download:") 
            
            
                    
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
