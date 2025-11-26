function exportOccurrenceProperties(modelName, outputFileName)
% exportOccurrenceProperties - Export occurrence numbers and properties to Excel
%
% This script searches through all components in a Systems Composer model,
% extracts OccurrenceNumber properties from stereotypes, and exports them
% to an Excel file with an embedded property matrix.
%
% Syntax:
%   exportOccurrenceProperties(modelName)
%   exportOccurrenceProperties(modelName, outputFileName)
%
% Inputs:
%   modelName      - Name of the Systems Composer model (without .slx extension)
%                    Default: 'NRC_Template'
%   outputFileName - Name of the output Excel file (without .xlsx extension)
%                    Default: 'OccurrenceProperties'
%
% Output Excel Format:
%   Column A: OccurrenceNumber - The occurrence number from stereotypes
%   Column B: PartNumber - Part number associated with the occurrence
%   Column C: ComponentName - Name(s) of components with this occurrence
%   Column D: Properties Matrix - Click to view detailed properties
%
% Example:
%   exportOccurrenceProperties('NRC_Template')
%   exportOccurrenceProperties('NRC_Template', 'MyOutput')
%
% The script creates two sheets:
%   1. Summary - Main table with occurrence numbers and component info
%   2. Properties - Detailed property breakdown for each occurrence

    % Get paths - save output to Tools folder
    scriptPath = fileparts(mfilename('fullpath'));
    toolsPath = fullfile(scriptPath, '..', '..', 'Tools');

    % Set default arguments
    if nargin < 1 || isempty(modelName)
        modelName = 'NRC_Template';
    end

    if nargin < 2 || isempty(outputFileName)
        outputFileName = 'OccurrenceProperties';
    end

    % Ensure .xlsx extension
    if ~endsWith(outputFileName, '.xlsx')
        outputFileName = [outputFileName '.xlsx'];
    end

    % Always save to Tools folder
    outputFileName = fullfile(toolsPath, outputFileName);

    fprintf('Extracting occurrence data from model: %s\n', modelName);

    try
        % Extract data using backend function
        componentData = extractOccurrenceData(modelName);

        if isempty(componentData)
            warning('No components with OccurrenceNumber property found in model: %s', modelName);
            return;
        end

        fprintf('Found %d unique occurrence numbers\n', length(componentData));

        % Create Excel file
        createExcelOutput(componentData, outputFileName);

        fprintf('Export complete! File saved as: %s\n', outputFileName);

    catch ME
        error('Failed to export occurrence properties:\n%s', ME.message);
    end

end

function createExcelOutput(componentData, fileName)
% Create Excel file with summary and detailed property sheets

    % Delete existing file if it exists
    if exist(fileName, 'file')
        delete(fileName);
    end

    % Prepare summary sheet data
    numRows = length(componentData);
    summaryData = cell(numRows + 1, 4);

    % Headers
    summaryData(1, :) = {'OccurrenceNumber', 'PartNumber', 'ComponentName', 'Properties'};

    % Fill in data rows
    for i = 1:numRows
        data = componentData(i);

        % Column A: Occurrence Number
        summaryData{i+1, 1} = data.OccurrenceNumber;

        % Column B: Part Number
        summaryData{i+1, 2} = data.PartNumber;

        % Column C: Component Names (join multiple names with semicolon)
        if length(data.ComponentNames) == 1
            summaryData{i+1, 3} = data.ComponentNames{1};
        else
            summaryData{i+1, 3} = strjoin(data.ComponentNames, '; ');
        end

        % Column D: Link to properties detail
        summaryData{i+1, 4} = sprintf('See Properties sheet (Row %d)', i+1);
    end

    % Write summary sheet
    writecell(summaryData, fileName, 'Sheet', 'Summary');

    % Create detailed properties sheet
    createPropertiesSheet(componentData, fileName);

    % Format the Excel file
    formatExcelFile(fileName, numRows);

end

function createPropertiesSheet(componentData, fileName)
% Create detailed properties sheet with all property information

    propSheetData = {'OccurrenceNumber', 'Component', 'Property Name', 'Property Value'};
    rowIdx = 2;

    for i = 1:length(componentData)
        data = componentData(i);
        occNum = data.OccurrenceNumber;

        % Process each component with this occurrence number
        for j = 1:length(data.Properties)
            compName = data.ComponentNames{j};
            props = data.Properties{j};

            % Get all property names and values
            if isstruct(props)
                propNames = fieldnames(props);

                for k = 1:length(propNames)
                    propName = propNames{k};
                    propValue = props.(propName);

                    % Convert value to string
                    if isnumeric(propValue)
                        propValueStr = num2str(propValue);
                    elseif ischar(propValue) || isstring(propValue)
                        propValueStr = char(propValue);
                    elseif islogical(propValue)
                        propValueStr = mat2str(propValue);
                    elseif iscell(propValue)
                        propValueStr = strjoin(cellfun(@num2str, propValue, 'UniformOutput', false), ', ');
                    else
                        try
                            propValueStr = char(propValue);
                        catch
                            propValueStr = class(propValue);
                        end
                    end

                    % Add row to sheet
                    propSheetData{rowIdx, 1} = occNum;
                    propSheetData{rowIdx, 2} = compName;
                    propSheetData{rowIdx, 3} = propName;
                    propSheetData{rowIdx, 4} = propValueStr;

                    rowIdx = rowIdx + 1;
                end
            end
        end
    end

    % Write properties sheet
    writecell(propSheetData, fileName, 'Sheet', 'Properties');

end

function formatExcelFile(fileName, numDataRows)
% Format the Excel file for better readability

    if ~ispc
        return; % Excel COM only works on Windows
    end

    Excel = [];
    try
        Excel = actxserver('Excel.Application');
        Excel.Visible = false;
        Excel.DisplayAlerts = false;

        % fileName is already an absolute path
        Workbook = Excel.Workbooks.Open(fileName);

        % Format Summary sheet
        Sheet = Workbook.Sheets.Item('Summary');
        Sheet.Activate;

        % Auto-fit columns
        Sheet.Columns.AutoFit;

        % Bold headers
        headerRange = Sheet.Range('A1:D1');
        headerRange.Font.Bold = true;
        headerRange.Interior.ColorIndex = 15; % Gray background

        % Add borders
        dataRange = Sheet.Range(sprintf('A1:D%d', numDataRows + 1));
        dataRange.Borders.LineStyle = 1;

        % Format Properties sheet
        Sheet = Workbook.Sheets.Item('Properties');
        Sheet.Activate;
        Sheet.Columns.AutoFit;

        headerRange = Sheet.Range('A1:D1');
        headerRange.Font.Bold = true;
        headerRange.Interior.ColorIndex = 15;

        % Save and close
        Workbook.Save;
        Workbook.Close;
        Excel.Quit;
        delete(Excel);

    catch ME
        % Clean up Excel if it was created
        try
            if ~isempty(Excel)
                Excel.Quit;
                delete(Excel);
            end
        catch
        end
        % Only warn if it's not a "file not found" type error
        if ~contains(ME.message, 'Could not find')
            warning('Could not apply Excel formatting: %s', ME.message);
        end
    end

end
