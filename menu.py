import pandas as pd
import os

# Load the data
file_path = 'data.xlsx'
df = pd.read_excel(file_path, sheet_name='The Data')

filtered_df = df[df.iloc[:, 0].str.contains(r'\bWine - ', na=False)]
# filtered_df = filtered_df.drop(12)

# Define LaTeX templates
def format_region(region):
    max_length = 20
    """Format the region to include a line break if it contains a city and region."""
    if "," in region:
        parts = region.split(",", 1)
        return f"{parts[0]}, \\\\{parts[1].strip()}"
    else:
        words = region.split()
        lines = []
        current_line = []

        for word in words:
            if sum(len(w) for w in current_line) + len(word) + len(current_line) <= max_length:
                current_line.append(word)
            else:
                lines.append(" ".join(current_line))
                current_line = [word]

        lines.append(" ".join(current_line))  # Add the last line
        return " \\\\ ".join(lines)
    
def format_grape_variety(grape_variety, max_length=40):
    """
    Format grape variety list: If the string exceeds max_length,
    split at a comma and push the remaining text to a new line.
    Trailing commas are preserved.

    Args:
        grape_variety (str): Input string with grape varieties.
        max_length (int): Maximum allowed length per line.

    Returns:
        str: Formatted string with '\\\\' added for new lines.
    """
    parts = grape_variety.split(",")  # Split at commas
    current_line = ""
    formatted_lines = []

    for i, part in enumerate(parts):
        part = part.strip()  # Remove extra spaces
        # part = part.strip('/n')
        # print("HERE")
        # print(part)
        addition = (part + ",") if i < len(parts) - 1 else part  # Add comma if not last part
        
        # Check if adding this part exceeds the max_length
        if len(current_line) + len(addition) + 1 <= max_length:  # +1 for the space
            current_line += (" " if current_line else "") + addition
        else:
            formatted_lines.append(current_line.strip())  # Add current line to output
            current_line = addition  # Start a new line with the current part

    if current_line:
        formatted_lines.append(current_line.strip())  # Add any remaining part

    return " \\\\\n".join(formatted_lines)  # Join lines with LaTeX newline

def generate_wine_entry(glass_price, bottle_price, vintage, winery, wine_name, grape_variety, winemaker, region, description):
    """
    Generate LaTeX code for a wine entry.
    
    Arguments:
    glass_price    - Price for a glass of wine
    bottle_price   - Price for a bottle of wine
    vintage        - Wine vintage (year)
    winery         - Winery name
    wine_name      - Name of the wine
    grape_variety  - Grape variety used
    winemaker      - Winemaker(s) name(s)
    region         - Wine region
    description    - Description of the wine
    
    Returns:
    LaTeX formatted string
    """
    return f"""    
    {{\\\\{glass_price} / {bottle_price}}} & {{{vintage} {winery} {wine_name} \\\\ {format_grape_variety(grape_variety)} \\\\ {winemaker}}} & {{{format_region(region)}}} \\\\
    \\\\
    \\SetCell[c=3]{{\\linewidth}}{{{description}}} \\\\
    \\SetCell[c=3]{{\\linewidth}} & & \\\\
"""

def generate_wine_menu(data):
    current_heading = None
    table = ""
    count = 0

    for index, row in data.iterrows():
        # Check for new heading
        if row["Heading"] != current_heading:
            # Close previous longtblr environment if necessary
            if current_heading is not None:
                table += "\\end{longtblr}\n\n\\vspace{-15pt} \n"
            # Start new longtblr environment
            current_heading = row["Heading"]
            count += 1
            table += f"""
\\begin{{longtblr}}[
    theme = TASMenu,
    caption = \\LARGE{{{current_heading}}},
    halign = j,
    valign = m,
]{{
    width = \\linewidth,
    colspec = llr,
}}

\\hline\\hline
    \\SetCell[c=3]{{\\linewidth}} & & \\\\
            """
            # count = 0  # Reset the count for the new heading

        # Prepare the row data
        glass_price = row["Glass"] if pd.notna(row["Glass"]) else ""
        glass_price = int(glass_price) if isinstance(glass_price, (float, int)) and glass_price == int(glass_price) else glass_price
        bottle_price = row["Bottle"] if pd.notna(row["Bottle"]) else ""
        bottle_price = int(bottle_price) if isinstance(bottle_price, (float, int)) and bottle_price == int(bottle_price) else bottle_price
        vintage = row["Vintage"] if pd.notna(row["Vintage"]) else ""
        winery = row["Winery"] if pd.notna(row["Winery"]) else ""
        wine_name = row["Name"] if pd.notna(row["Name"]) else ""
        grape_variety = row["Grape Variety"] if pd.notna(row["Grape Variety"]) else ""
        winemaker = row["Winemaker and/or Owner"] if pd.notna(row["Winemaker and/or Owner"]) else ""
        region = row["Region"] if pd.notna(row["Region"]) else ""
        description = row["Description"] if pd.notna(row["Description"]) else ""

        # Add the wine entry
        table += generate_wine_entry(glass_price, bottle_price, vintage, winery, wine_name, grape_variety, winemaker, region, description)
        count += 1

        # Add a pagebreak if needed
        if count >= 4:
            table += """    \\pagebreak
    \\\\"""
            count = 0

    # Close the last longtblr environment
    if current_heading is not None:
        table += "\\end{longtblr}\n\n\\vspace{-15pt} \n"
    
    table = table.replace(
        """    \\pagebreak
    \\\\\\end{longtblr}

\\vspace{-15pt} 
""",
"""    
    \\end{longtblr}

\\vspace{-15pt} 
\\pagebreak"""
    )

    table = table.replace(
        """
    {\\17 / 84} & {2021 Topper's Mountain Wines "Hill of Dreams" \\ Sauvignon Blanc, \\
Verdejo Grüner Veltliner, \\ Mark and Stephanie Kirkby} & {New England} \\
    \\
""","""
    {\\17 / 84} & {2021 Topper's Mountain Wines "Hill of Dreams" \\ Sauvignon Blanc, Verdejo, Grüner Veltliner \\ Mark and Stephanie Kirkby} & {New England} \\
    \\
""")

    return table


# Save the LaTeX output to a file or print
with open("wine_tables.tex", "w") as file:
    file.write(generate_wine_menu(filtered_df))

file.close()

# print(latex_output) 

### Beer Processing

beer_cider_df = df[df["Heading"].str.contains(r'\bBeer & Cider\b', na=False)]

def format_beer_cider_row(row):
    """
    Generate LaTeX row for Beer & Cider table.
    Maps `Winery`, `Region`, and `Name` to the LaTeX format.
    """
    winery = row["Winery"] if pd.notna(row["Winery"]) else ""
    region = row["Region"] if pd.notna(row["Region"]) else ""
    name = row["Name"] if pd.notna(row["Name"]) else ""
    glass_price = f"{int(row['Glass'])}" if row["Glass"].is_integer() else f"{row['Glass']:.1f}" if pd.notna(row["Glass"]) else ""

    
    return rf"""
\SetCell[c=3]{{\linewidth}} & & \\
{glass_price} & {{{winery} \\ {region}}} & {name} \\
"""

# Function to generate the LaTeX table
def generate_beer_cider_table(data):
    table = r"""
\begin{longtblr}[
    theme = TASMenu,
    caption = \LARGE{Beer \& Cider},
    halign = j,
    valign = m,
]{
    width = \textwidth,
    colspec = cll,
    % hlines,
    % vlines,
}
\hline\hline
"""
    # Generate rows
    for _, row in data.iterrows():
        table += format_beer_cider_row(row)

    # Close the table
    table += r"""
\end{longtblr}
"""
    return table


# Generate the LaTeX table
latex_table = generate_beer_cider_table(beer_cider_df)

latex_table = latex_table.replace("""Ask about our latest tap beers, \\\\ or just look at our tap.""", """{Ask about our latest tap beers, \\\\ or just look at our tap.}""")

# Save the LaTeX output to a file
with open("beer_cider_table.tex", "w") as file:
    file.write(latex_table)

file.close()

# print(latex_table)

### Cocktail Processing

cocktails_df = df[df["Heading"].str.contains(r'\b(Cocktails|Mocktail)\b', na=False)]


# Important for some names
def escape_latex(text):
    """
    Escapes special LaTeX characters in the text.
    """
    if not isinstance(text, str):
        return text  # If not a string, return as-is

    # Dictionary of LaTeX special characters and their escaped versions
    latex_special_chars = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '^': r'\textasciicircum{}',
    }
    
    # Replace each special character with its escaped version
    for char, escape in latex_special_chars.items():
        text = text.replace(char, escape)
    
    return text

# Function to format LaTeX rows for cocktails
def format_cocktail_row(row):
    """
    Generate LaTeX row for Cocktails table.
    Maps `Glass`, `Name`, and `Description` to the LaTeX format.
    """
    glass_price = int(row["Glass"]) if pd.notna(row["Glass"]) else ""
    name = row["Name"] if pd.notna(row["Name"]) else ""
    name = escape_latex(name)
    description = row["Grape Variety"] if pd.notna(row["Grape Variety"]) else ""

    return rf"""
    {glass_price} & {name} & {description} \\
    \SetCell[c=3]{{\linewidth}} & & \\
"""

# Function to generate the LaTeX table
def generate_cocktail_table(data):
    table = r"""
\begin{longtblr}[
    theme = TASMenu,
    caption = \LARGE{Cocktails},
    halign = j,
    valign = m,
]{
    width = \linewidth,
    colspec = cll,
    % hlines,
    % vlines,
}
\hline\hline
    \SetCell[c=3]{\linewidth} & & \\
"""
    # Generate rows
    for _, row in data.iterrows():
        table += format_cocktail_row(row)

    # Close the table
    table += r"""
\end{longtblr}
"""
    return table


latex_table = generate_cocktail_table(cocktails_df)

# Save the LaTeX output to a file
with open("cocktails_table.tex", "w") as file:
    file.write(latex_table)

# print(latex_table)



### Spirits Processing

spirits_categories = ["Gin", "Vodka", "Whisky", "Rum", "Liqueur"]

# Function to format LaTeX rows for spirits
def format_spirit_row(row, displayed_names=set()):
    """
    Generate LaTeX row for spirits table.
    Escapes special characters in `Name` and `Grape Variety`.
    """
    glass_price = f"{int(row['Glass'])}" if row["Glass"].is_integer() else f"{row['Glass']:.1f}" if pd.notna(row["Glass"]) else ""
    name = escape_latex(row["Name"]) if pd.notna(row["Name"]) else ""
    grape_variety = escape_latex(row["Grape Variety"]) if pd.notna(row["Grape Variety"]) else ""
    region = escape_latex(row["Region"]) if pd.notna(row["Region"]) else ""

    combined_name_location = f"{name} \\\\ {region}" if region else name
    if combined_name_location in displayed_names:
        combined_name_location = ""  # Replace with an empty string if repeated
    else:
        displayed_names.add(combined_name_location)  # Mark as displayed

    return (rf"""
    {glass_price} & {{{combined_name_location}}} & {{{grape_variety}}} \\
    \SetCell[c=3]{{\linewidth}} & & \\
""", displayed_names)


# Function to generate the LaTeX table for each spirit
def generate_spirit_table(data, spirit_name):
    table = rf"""
\needspace{{10\baselineskip}}
\begin{{longtblr}}[
    theme = TASMenu,
    caption = \LARGE{{Spirits - {spirit_name}}},
    halign = j,
    valign = m,
]{{
    width = \linewidth,
    colspec = cll,
    % hlines,
    % vlines,
}}
\hline\hline
    \SetCell[c=3]{{\linewidth}} & & \\
"""
    disp_names = set()  # Track displayed names to avoid repetition
    # Generate rows
    for _, row in data.iterrows():
        res = format_spirit_row(row, displayed_names=disp_names)
        table += res[0]
        disp_names = res[1]

    # Close the table
    table += r"""
\end{longtblr}
"""
    return table


latex_output = ""
for spirit in spirits_categories:
    spirit_df = df[df["Heading"].str.contains(fr'\b{spirit}\b', na=False)]  # Filter rows for the current spirit
    if not spirit_df.empty:
        latex_output += generate_spirit_table(spirit_df, spirit) + "\n"

        
with open("spirits_tables.tex", "w") as file:
    file.write(latex_output)

# print(latex_output)



### More Spirits Processing

# Function to format LaTeX rows for spirits
def format_more_spirit_row(row):
    """
    Generate LaTeX row for spirits table.
    Escapes special characters in `Name` and `Grape Variety`.
    """
    glass_price = f"{int(row['Glass'])}" if row["Glass"].is_integer() else f"{row['Glass']:.1f}" if pd.notna(row["Glass"]) else ""
    name = escape_latex(row["Name"]) if pd.notna(row["Name"]) else ""
    grape_variety = escape_latex(row["Grape Variety"]) if pd.notna(row["Grape Variety"]) else ""
    region = escape_latex(row["Region"]) if pd.notna(row["Region"]) else ""

    combined_name_location = f"{name} \\\\ {region}" if region else name

    return rf"""
    {glass_price} & {{{combined_name_location}}} & {{{grape_variety}}} \\
    \SetCell[c=3]{{\linewidth}} & & \\
"""

def generate_more_spirits_table(data):
    """
    Generate LaTeX table for Pisco, Soju, Amaro, Vermouth, or PX under the title 'More Spirits from NSW'.
    """
    table = r"""
\begin{longtblr}[
    theme = TASMenu,
    caption = \LARGE{More Spirits from NSW},
    halign = j,
    valign = m,
]{
    width = \linewidth,
    colspec = cll,
    % hlines,
    % vlines,
}
\hline\hline
    \SetCell[c=3]{\linewidth} & & \\

"""
    # Initialize a set to track displayed names
    displayed_names = set()

    # Generate rows
    for _, row in data.iterrows():
        table += format_more_spirit_row(row)

    # Close the table
    table += r"""
\end{longtblr}
"""
    return table

more_spirits_keywords = ["Pisco", "Soju", "Amaro", "Vermouth", "PX"]
pattern = "|".join(more_spirits_keywords)  # Create regex pattern for these keywords
more_spirits_df = df[df["Heading"].str.contains(pattern, na=False, case=False)]

# Generate LaTeX table for More Spirits from NSW
more_spirits_latex = generate_more_spirits_table(more_spirits_df)

# Save the LaTeX output to a file
with open("more_spirits_table.tex", "w") as file:
    file.write(more_spirits_latex)

# print(more_spirits_latex)


### Non-Alcoholic Processing


def generate_non_alcoholic_table(data):
    """
    Generate LaTeX table for Non-alcoholic items.
    """
    table = r"""
\begin{longtblr}[
    theme = TASMenu,
    caption = \LARGE{Non-alcoholic},
    halign = j,
    valign = m,
]{
    width = \linewidth,
    colspec = cll,
    % hlines,
    % vlines,
}
\hline\hline\\
"""
    # Generate rows
    for _, row in data.iterrows():
        glass_price = (
            f"{int(row['Glass'])}" if pd.notna(row["Glass"]) and row["Glass"].is_integer()
            else f"{row['Glass']:.1f}" if pd.notna(row["Glass"])
            else "~"
        )
        name = escape_latex(row["Name"]) if pd.notna(row["Name"]) else "~"
        region = escape_latex(row["Region"]) if pd.notna(row["Region"]) else ""
        grape_variety = escape_latex(row["Grape Variety"]) if pd.notna(row["Grape Variety"]) else "~"
    
        # Combine name and location, deduplicate if repeated
        combined_name_location = f"{name} \\\\ {region}" if region else name

        # Add rows to the table
        table += rf"""
    {glass_price} & {{{combined_name_location}}} & {grape_variety} \\
    \SetCell[c=3]{{\linewidth}} & & \\
"""
        
    table += r"""
    4.5 & Sparkling water & Unlimited  \\
    \SetCell[c=3]{\linewidth} & & \\
    \vspace{-10pt}
    -   & {tap\textsuperscript{\texttrademark} by Sydney Water \\ Wollondilly Shire } & ~ \\
    \SetCell[c=3]{\linewidth, halign=l} Bearing no notes or hints of anything, this special blend suits all tastes. Officially known as ``tap\textsuperscript{\texttrademark} A Sydney Water Product'', locals refer to it as the ``Warragamba Slammer'' & ~ & ~ \\
    """

    table += r"""
\end{longtblr}
"""
    return table

# Filter rows for 'Non-alcoholic'
non_alcoholic_df = df[df["Heading"].str.contains(r'\bNon-alcoholic\b', na=False, case=False)]
non_alcoholic_df = non_alcoholic_df.iloc[1:-1]

# Generate LaTeX table for Non-alcoholic items
non_alcoholic_latex = generate_non_alcoholic_table(non_alcoholic_df)

# Save the LaTeX output to a file
with open("non_alcoholic_table.tex", "w") as file:
    file.write(non_alcoholic_latex)

# print(non_alcoholic_latex)


### Food Processing

def generate_food_table(data):
    """
    Generate LaTeX table for food items.
    """
    table = r"""
\begin{longtblr}[
    theme = TASMenu,
    caption = \LARGE{Food},
    halign = j,
    valign = m,
]{
    width = \linewidth,
    colspec = cll,
    % hlines,
    % vlines,
}
\hline\hline
\\
"""
    # Generate rows
    for _, row in data.iterrows():
        glass_price = (
            f"{int(row['Glass'])}" if pd.notna(row["Glass"]) and row["Glass"].is_integer()
            else f"{row['Glass']:.1f}" if pd.notna(row["Glass"])
            else "~"
        )
        name = escape_latex(row["Name"]) if pd.notna(row["Name"]) else "~"
        description = escape_latex(row["Grape Variety"]) if pd.notna(row["Grape Variety"]) else "~"

        # Add row to the table
        table += rf"""
    {glass_price} & {name} & {{{description}}} \\
    \SetCell[c=3]{{\linewidth}} & & \\
"""
    table += r"""
\end{longtblr}
"""
    return table

# Filter rows for food
food_df = df[df["Heading"].str.contains(r'\bFood\b', na=False, case=False)]

# Generate LaTeX table for food items
food_latex = generate_food_table(food_df)

# Save the LaTeX output to a file
with open("food_table.tex", "w") as file:
    file.write(food_latex)

print(food_latex)