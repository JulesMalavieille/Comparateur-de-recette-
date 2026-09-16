"""
Created on Tue Sep 15 18:42:45 2026

@author: Jules Malavieille
"""
    
    
import streamlit as st
import pandas as pd 
import io 


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

def split_ingredients(text):
    ingredients = []
    element = ""
    depth = 0

    for char in text:

        if char == "(":
            depth += 1

        elif char == ")":
            depth -= 1

        if char == "," and depth == 0:
            ingredients.append(element.strip())
            element = ""

        else:
            element += char

    if element:
        ingredients.append(element.strip())

    return ingredients



def compare_recipes(textA, textB):

    A = split_ingredients(textA)
    B = split_ingredients(textB)

    removed = [ingredient for ingredient in A if ingredient not in B]

    added = [ingredient for ingredient in B if ingredient not in A]

    unchanged = [ingredient for ingredient in A if ingredient in B]

    return removed, added, unchanged


def create_excel(textA, textB, removed, added, unchanged):

    max_len = max(len(removed), len(added), len(unchanged))

    removed_col = removed + [""] * (max_len - len(removed))
    added_col = added + [""] * (max_len - len(added))
    unchanged_col = unchanged + [""] * (max_len - len(unchanged))

    df_results = pd.DataFrame({
        "Retiré": removed_col,
        "Ajouté": added_col,
        "Inchangé": unchanged_col
    })

    df_recipes = pd.DataFrame({
        "Recette A": [textA],
        "Recette B": [textB]
    })

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        df_results.to_excel(
            writer,
            sheet_name="Comparaison",
            index=False
        )

        df_recipes.to_excel(
            writer,
            sheet_name="Recettes",
            index=False
        )

        worksheet_results = writer.sheets["Comparaison"]
        worksheet_results.set_column("A:C", 50)

        worksheet_recipes = writer.sheets["Recettes"]
        worksheet_recipes.set_column("A:B", 80)

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
# ZONES DE TEXTE
# =========================================================

colA, colB = st.columns(2, gap="large")


with colA:

    with st.container(border=True):

        st.markdown("### Recette A")
        st.caption("Version initiale")

        textA = st.text_area(
            "Recette A",
            height=250,
            placeholder="Collez la première liste d'ingrédients...",
            label_visibility="collapsed", 
            key="recipe_a")


with colB:

    with st.container(border=True):

        st.markdown("### Recette B")
        st.caption("Nouvelle version")

        textB = st.text_area(
            "Recette B",
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
            textA,
            textB,
            removed,
            added,
            unchanged)

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
            
            
                    
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            