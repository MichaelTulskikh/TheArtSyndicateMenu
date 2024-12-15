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


def wine_template(row):
    """Generate LaTeX for a wine item."""
    formatted_region = format_region(row['Region'])
    return rf"""
    \LineItem{{{row['Glass']:.0f}}}{{{row['Bottle']:.0f}}}{{{row['Vintage']}}}{{{row['Winery']}}}{{{row['Name'] if pd.notna(row['Name']) else ''}}}{{\\{row['Grape Variety']}}}{{{row['Winemaker and/or Owner']}}}{{{formatted_region}}}
    \BlankRow
    \ItemDescription{{{row['Description']}}}
    \BlankRow
    """

def wine_template_row_end(row):
    """Generate LaTeX for a wine item."""
    formatted_region = format_region(row['Region'])
    return rf"""
    \LineItem{{{row['Glass']:.0f}}}{{{row['Bottle']:.0f}}}{{{row['Vintage']}}}{{{row['Winery']}}}{{{row['Name'] if pd.notna(row['Name']) else ''}}}{{\\{row['Grape Variety']}}}{{{row['Winemaker and/or Owner']}}}{{{formatted_region}}}
    \BlankRow
    \ItemDescription{{{row['Description']}}}
    """

def generate_table(data, previous_heading=None):
    table = "\\TableStart\n"
    current_heading = None
    prev_heading = previous_heading
    for _, row in data.iterrows():
        print(row["Heading"], current_heading)
        if row["Heading"] != current_heading:
            current_heading = row["Heading"]
            if _ % 4 == 0 and current_heading == prev_heading:
                # Add (continued) if the first row's heading matches the previous chunk's last heading
                table += f"    \\MenuSection{{{current_heading} (continued)}}\n\n"
            else:
                table += f"    \\MenuSection{{{current_heading}}}\n\n"
        

        winery = row["Winery"] if pd.notna(row["Winery"]) else ""
        name = row["Name"] if pd.notna(row["Name"]) else ""
        grape_variety = row["Grape Variety"] if pd.notna(row["Grape Variety"]) else ""
        winemaker = row["Winemaker and/or Owner"] if pd.notna(row["Winemaker and/or Owner"]) else ""
        region = row["Region"] if pd.notna(row["Region"]) else ""
        description = row["Description"] if pd.notna(row["Description"]) else ""

        # Add line item
        table += (
            f"    \\LineItem{{{int(row['Glass'])}}}{{{int(row['Bottle'])}}}{{{row['Vintage']}}}"
            f"{{{winery}}}{{{name}}}{{\\\\{grape_variety}}}"
            f"{{{winemaker}}}{{{format_region(region)}}}\n"
        )
        table += "    \\BlankRow\n"
        
        # Add description
        table += f"    \\ItemDescription{{{row['Description']}}}\n"
        table += "    \\BlankRow\n"
    
    print('-----')
    table += "\\TableEnd\n"
    return [table, current_heading]
    

# Define the chunk size
chunk_size = 4

# Group into chunks of 4 rows each
chunks = [filtered_df.iloc[i:i + chunk_size] for i in range(0, len(filtered_df), chunk_size)]

# # Display chunks for verification
# for idx, chunk in enumerate(chunks):
#     print(f"Chunk {idx + 1}:\n", chunk, "\n")

# Generate LaTeX tables for all chunks
latex_output = ""
prev_heading = None
for i, chunk in enumerate(chunks, start=1):
    latex_output += f"% Chunk {i}\n"
    res = generate_table(chunk, prev_heading)
    latex_output += res[0]
    prev_heading = res[1]
    # print(prev_heading)
    latex_output += "\n"  # Separate chunks with a newline

# Save the LaTeX output to a file or print
with open("wine_tables.tex", "w") as file:
    file.write(latex_output)

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
    glass_price = int(row["Glass"]) if pd.notna(row["Glass"]) else ""
    
    return rf"""
\SetCell[c=3]{{\linewidth}} & & \\
{glass_price} & {{{winery} \\ {region}}} & ``{name}'' \\
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
    4   & Coffee: Espresso    & Numero Uno Coffee Roasters, St Peters.\\
    \SetCell[c=3]{\linewidth} & & \\

    5   & Coffee: Long black  & \\
    \SetCell[c=3]{\linewidth} & & \\

    5.5 & Coffee: White       & \\
    \SetCell[c=3]{\linewidth} & & \\
    \\
    \\
    \\
    \\
    4.5 & Sparkling water & Unlimited refills \\
    \SetCell[c=3]{\linewidth} & & \\

    -   & {tap\textsuperscript{\texttrademark} by Sydney Water \\ Wollondilly Shire } & ~ \\
    \SetCell[c=3]{\linewidth, halign=l} Bearing no notes or hints of anything, this special blend suits all tastes. Officially known as ``tap\textsuperscript{\texttrademark} A Sydney Water Product'', locals refer to it as the ``Warragamba Slammer'' & ~ & ~ \\
    """

    table += r"""
\end{longtblr}
"""
    return table

# Filter rows for 'Non-alcoholic'
non_alcoholic_df = df[df["Heading"].str.contains(r'\bNon-alcoholic\b', na=False, case=False)]
non_alcoholic_df = non_alcoholic_df.iloc[:-5]

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