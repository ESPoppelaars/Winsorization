%% Clear workspace and command window.
clear
clc

%% Instructions.
% This script identifies outliers and winsorizes them.
% Created by Eefje Poppelaars, May 2018.

% FIRST.
% Use Excel to replace missing values with NaN!
% Create an Excel file, containing only the subject numbers in the first
% colum, and data you want to winsorize in the remaining colums.

% ENTER HERE.
% Enter an Excelfile name, then run the script.
filename = '20180502_IJMData_Carma.xlsx';

%% Import data.
%Read the numeric data and the text headers from the excel file.
[data,headers] = xlsread(filename);
headers = headers(1,:);

% Create subject indices.
    % Sample size.
    ppnnum = size(data,1);
    % List of subject indices.
    subjects = [1:ppnnum]';

% Initialize winsorized data to be exported later.
winsor = data;
% Initialize count of outliers per variable.
count = zeros(1,size(data,2));

%% Identify and winsorize outliers.
% Identify outliers as more extreme than (-)3.3 SDs, 
% and winsorize outliers by replacing them with the nearest non-extreme
% value + 1%.
for i = 2:size(data,2)
    % Ignore the 'not a number' values.
        % Subject numbers without NaN.
        nonNaN = subjects( ~isnan(data(:,i)) );
        % Data without NaN.
        dataRaw = data(nonNaN,i);
    % For every colum, create new variable: z-transform.
    Z = zscore(dataRaw);
    % Combine the subject indices with the Z-scores.
    Zindexed = [nonNaN, Z];
    
    % Run through Z to find the highest and lowest non-extreme value.
    nnnHigh = 0;
    nnnLow = 0;
    for j = 1:size(Zindexed,1)
        if Zindexed(j,2) > nnnHigh && Zindexed(j,2) < 3.3
            nnnHigh = Zindexed(j,2);
            nnnHighRaw = dataRaw(j);
        end
        if Zindexed(j,2) < nnnLow && Zindexed(j,2) > -3.3
            nnnLow = Zindexed(j,2);
            nnnLowRaw = dataRaw(j);
        end
    end
    
    %   For every value that is > 3.3 or < -3.3, replace the original value
    %   with the nearest non-extreme value + 1%.
    %   For every replacement, add 1 to the counter.
    for j = 1:size(Zindexed,1)
        if Zindexed(j,2) > 3.3
            winsor(Zindexed(j,1),i) = nnnHighRaw*1.01;
            count(1,i) = count(1,i)+1;
        elseif Zindexed(j,2) < -3.3
            winsor(Zindexed(j,1),i) = nnnLowRaw*0.99;
            count(1,i) = count(1,i)+1;
        end
    end
end

%% Export.
% Add 'Winsor_' to the headers.
% Concatonate headers with data for export.
exportfilename = strcat('Winsor_',filename);
export = {};
winsor_headers = strcat('Winsor_',headers);
for i = 1:length(winsor_headers)
    export(1,i) = winsor_headers(i);
    for j = 1:size(winsor,1)
        export(j+1,i) = {winsor(j,i)};
    end
end

% Add 'Count_' to the headers.
% Concatonate headers with count for export.
exportfilename_count = strcat('Count_',filename);
export_count = {};
count_headers = strcat('Count_',headers);
for i = 1:length(count_headers)
    export_count(1,i) = count_headers(i);
    for j = 1:size(count,1)
        export_count(j+1,i) = {count(j,i)};
    end
end

% Write excel files.
% Winsorized data.
xlswrite(exportfilename,export);
% Count of winsorized outliers per variable.
xlswrite(exportfilename_count,export_count);