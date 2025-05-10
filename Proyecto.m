%% Leer archivo Excel
archivo = 'Consumos 2023.xlsx'; % asegúrate de tener este archivo
opts = detectImportOptions(archivo, 'Sheet', 'Informe 1', 'NumHeaderLines', 4);
opts.VariableNamingRule = 'preserve';
data = readtable(archivo, opts);

%% Calcular consumo total anual (Kwh N1 a N12)
columnas_kwh = {};
for i = 1:12
    colname = sprintf('Kwh N%d', i);
    if ismember(colname, data.Properties.VariableNames)
        columnas_kwh{end+1} = colname;
    end
end

consumo_total = zeros(height(data), 1);
for i = 1:length(columnas_kwh)
    columna = data.(columnas_kwh{i});
    if iscell(columna)
        columna = str2double(columna);
    end
    columna(isnan(columna)) = 0;
    consumo_total = consumo_total + columna;
end
data.ConsumoTotalAnual = consumo_total;

%% Agrupar por parroquia
[grupos, parroquias] = findgroups(data.Parroquia);
consumo_por_parroquia = splitapply(@sum, data.ConsumoTotalAnual, grupos);

tabla_consumo = table(parroquias, consumo_por_parroquia, ...
    'VariableNames', {'Parroquia', 'ConsumoTotalAnual'});
tabla_consumo = sortrows(tabla_consumo, 'ConsumoTotalAnual', 'descend');

%% Estadística descriptiva
media_c = mean(tabla_consumo.ConsumoTotalAnual);
mediana_c = median(tabla_consumo.ConsumoTotalAnual);
std_c = std(tabla_consumo.ConsumoTotalAnual);
p25 = prctile(tabla_consumo.ConsumoTotalAnual, 25);
p75 = prctile(tabla_consumo.ConsumoTotalAnual, 75);

%% Clasificación por percentiles
clasificacion = repmat("Medio", height(tabla_consumo), 1);
clasificacion(tabla_consumo.ConsumoTotalAnual <= p25) = "Bajo";
clasificacion(tabla_consumo.ConsumoTotalAnual >= p75) = "Alto";
tabla_consumo.Clasificacion = clasificacion;

%% Regla empírica y Chebyshev
z_scores = (tabla_consumo.ConsumoTotalAnual - media_c) / std_c;
emp1 = sum(abs(z_scores) <= 1) / height(tabla_consumo) * 100;
emp2 = sum(abs(z_scores) <= 2) / height(tabla_consumo) * 100;
emp3 = sum(abs(z_scores) <= 3) / height(tabla_consumo) * 100;
chebyshev_2 = 1 - 1 / (2^2);

%% Probabilidad de superar cierto umbral
umbral = 20000;
P_umbral = sum(tabla_consumo.ConsumoTotalAnual > umbral) / height(tabla_consumo) * 100;

%% Combinaciones de parroquias con consumo alto
parr_altas = tabla_consumo.Parroquia(tabla_consumo.ConsumoTotalAnual > media_c);
n_altas = length(parr_altas);
if n_altas >= 3
    combinaciones3 = nchoosek(n_altas, 3);
else
    combinaciones3 = 0;
end

%% Mostrar resultados
fprintf('\n--- Análisis Estadístico y Probabilístico ---\n');
fprintf('Media: %.2f kWh\n', media_c);
fprintf('Mediana: %.2f kWh\n', mediana_c);
fprintf('Desviación estándar: %.2f kWh\n', std_c);
fprintf('Percentil 25: %.2f\n', p25);
fprintf('Percentil 75: %.2f\n', p75);
fprintf('Parroquias > media: %d\n', n_altas);
fprintf('Combinaciones posibles (3): %d\n', combinaciones3);
fprintf('Probabilidad > %.0f kWh: %.2f%%\n', umbral, P_umbral);
fprintf('Regla empírica (1σ): %.2f%%\n', emp1);
fprintf('Regla empírica (2σ): %.2f%%\n', emp2);
fprintf('Regla empírica (3σ): %.2f%%\n', emp3);
fprintf('Límite Chebyshev k=2: %.2f%%\n', chebyshev_2 * 100);

%% Gráfico de barras con líneas de referencia
figure('Position', [100 100 1000 500]);
bar(categorical(tabla_consumo.Parroquia), tabla_consumo.ConsumoTotalAnual, ...
    'FaceColor', [0.2 0.8 0.6], 'EdgeColor', 'black');
hold on;
yline(media_c, '--r', sprintf('Media: %.0f', media_c));
yline(mediana_c, '--g', sprintf('Mediana: %.0f', mediana_c));
hold off;
title('Consumo Total Anual por Parroquia');
xlabel('Parroquia');
ylabel('Consumo (kWh)');
xtickangle(45);
grid on;


