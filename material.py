import pandas as pd

def calculate_results(data):
    # Create a copy of the data
    data = data.copy()

    data['solids_oven_average'] = (data['solids_oven1'] + data['solids_oven2']) / 2
    data['deviation'] = abs(data['solids_oven_average'] - data['solids_smart'])

    grouped = data.groupby(['material_number', 'method'])[['deviation', 'drying_time', 'ew']].agg({
        'deviation': ['count', 'mean', 'std'],
        'drying_time': 'mean',
        'ew': 'mean'
    })

    results_list = []

    for (material_number, method), row in grouped.iterrows():
        material = data[data['material_number'] == material_number]['material'].mode().iloc[0]
        deviation_values = ', '.join(map("{:.2f}".format, data[(data['material_number'] == material_number) & (data['method'] == method)]['deviation']))

        results_list.append({
            'material_number': material_number,
            'material': material,
            'method': method,
            'count': row[('deviation', 'count')],
            'avg_deviation': row[('deviation', 'mean')],
            'deviation_values': deviation_values
        })

    return pd.DataFrame(results_list)

def get_material_data(material_number, data):
    # Try to convert the material_number to integer, if possible
    try:
        material_number = int(material_number)
    except ValueError:
        pass  # It's a string, no conversion needed

    filtered_data = data[data['material_number'] == material_number]
    
    if filtered_data.empty:
        print(f"No data found for material number: {material_number}")
        print("Here are the material numbers in the dataset:")
        print(data['material_number'].unique())  # Print unique material numbers in the dataset for debugging
        return

    results = calculate_results(filtered_data)
    
    # Get the material name
    material_name = results['material'].iloc[0]
    print(f"Material: {material_name}\n")
    
    # Find the method with the lowest avg_deviation
    best_method = results['avg_deviation'].idxmin()

    # Table header
    print("{:<15} | {:<6} | {:<13} | {:<12} | {}".format("Method", "Count", "Avg Deviation", "Best Method", "Deviation Values"))
    print("-" * 80)  # Adjusted line length
    
    for index, row in results.iterrows():
        best_mark = 'x' if index == best_method else ''
        print("{:<15} | {:<6} | {:<13.2f} | {:<12} | {}".format(row['method'], row['count'], row['avg_deviation'], best_mark, row['deviation_values']))
    print("\n")



if __name__ == '__main__':
    data = pd.read_excel('smart6.xlsx', engine='openpyxl')
    
    while True:  # Keep prompting for input until the loop is broken
        material_number_input = input("Please enter a material number (or 'x' to exit): ")
        
        if material_number_input.lower() == 'x':
            break
        
        get_material_data(material_number_input, data)

