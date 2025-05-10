% Cargar archivo Excel
archivo = 'BIL038_-_Histórico_de_Consumos 2022.xlsx';
opts = detectImportOptions(archivo, 'Sheet', 'Informe 1', 'NumHeaderLines', 5);
df = readtable(archivo, opts);

% Filtrar columnas que comienzan con 'Kwh N' seguido de un número entre 1 y 12
colNames = df.Properties.VariableNames;
columnas_kwh = colNames(contains(colNames, 'KwhN'));

% Asegurar que solo tome hasta 'Kwh N12'
columnas_kwh = columnas_kwh(cellfun(@(x) ...
    all(isstrprop(x(end), 'digit')) && str2double(x(end)) <= 12, columnas_kwh));

% Calcular consumo total anual
df.ConsumoTotalAnual = sum(df{:, columnas_kwh}, 2, 'omitnan');

% Seleccionar columnas necesarias
df_sector = df(:, {'Parroquia', 'ConsumoTotalAnual'});

% Agrupar por parroquia y sumar
tabla_consumo_sector = groupsummary(df_sector, 'Parroquia', 'sum', 'ConsumoTotalAnual');

% Ordenar de mayor a menor
tabla_consumo_sector = sortrows(tabla_consumo_sector, 'sum_ConsumoTotalAnual', 'descend');

% Mostrar resultado
disp(tabla_consumo_sector)
