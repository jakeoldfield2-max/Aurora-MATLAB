function massBreakdown(modelName)
%MASSBREAKDOWN Generate a mass breakdown report for all components
%
%   massBreakdown() - Uses default model 'NRC_Template'
%   massBreakdown(modelName) - Uses specified model
%
%   This script:
%   1. Updates occurrence properties by running exportOccurrenceProperties
%   2. Creates a Mass Breakdown Excel sheet showing:
%      - OccurrenceNumber
%      - Component Name
%      - Mass
%      - Total mass of the rocket

    if nargin < 1
        modelName = 'NRC_Template';
    end

    % Get paths - save output to Tools folder
    scriptPath = fileparts(mfilename('fullpath'));
    toolsPath = fullfile(scriptPath, '..', '..', 'Tools');

    fprintf('=== Mass Breakdown Report ===\n\n');

    % Step 1: Update occurrence properties
    fprintf('Step 1: Updating occurrence properties...\n');
    exportOccurrenceProperties(modelName);
    fprintf('\n');

    % Step 2: Extract data directly for processing
    fprintf('Step 2: Generating mass breakdown...\n');
    componentData = extractOccurrenceData(modelName);

    if isempty(componentData)
        warning('No components with OccurrenceNumber found.');
        return;
    end

    % Step 3: Build mass breakdown table - save to Tools folder
    outputFile = fullfile(toolsPath, 'MassBreakdown.xlsx');

    % Prepare data
    tableData = {};
    totalMass = 0;
    rowIdx = 1;

    for i = 1:length(componentData)
        occNum = componentData(i).OccurrenceNumber;
        compNames = componentData(i).ComponentNames;
        props = componentData(i).Properties;

        % Get mass from properties (check each component's properties)
        for j = 1:length(props)
            compProps = props{j};
            mass = 0;

            if isstruct(compProps) && isfield(compProps, 'Mass')
                massVal = compProps.Mass;
                if isnumeric(massVal)
                    mass = massVal;
                elseif ischar(massVal) || isstring(massVal)
                    mass = str2double(massVal);
                    if isnan(mass), mass = 0; end
                end
            end

            compName = compNames{j};
            tableData{rowIdx, 1} = occNum;
            tableData{rowIdx, 2} = compName;
            tableData{rowIdx, 3} = mass;
            totalMass = totalMass + mass;
            rowIdx = rowIdx + 1;
        end
    end

    % Create Excel output
    if exist(outputFile, 'file')
        delete(outputFile);
    end

    % Headers
    headers = {'OccurrenceNumber', 'ComponentName', 'Mass (kg)'};

    % Add total row
    tableData{rowIdx, 1} = '';
    tableData{rowIdx, 2} = 'TOTAL';
    tableData{rowIdx, 3} = totalMass;

    % Combine headers and data
    outputData = [headers; tableData];

    % Write to Excel
    writecell(outputData, outputFile, 'Sheet', 'Mass Breakdown');

    % Format Excel (Windows only)
    formatBreakdownExcel(outputFile, 'Mass Breakdown', rowIdx);

    fprintf('\nMass Breakdown complete!\n');
    fprintf('Total Mass: %.2f kg\n', totalMass);
    fprintf('File saved: %s\n', outputFile);
end

function formatBreakdownExcel(fileName, sheetName, numDataRows)
    try
        if ispc
            Excel = actxserver('Excel.Application');
            Excel.Visible = false;
            Workbook = Excel.Workbooks.Open(fileName);
            Sheet = Workbook.Sheets.Item(sheetName);
            Sheet.Activate;

            % Auto-fit columns
            Sheet.Columns.AutoFit;

            % Bold headers
            headerRange = Sheet.Range('A1:C1');
            headerRange.Font.Bold = true;
            headerRange.Interior.ColorIndex = 15;

            % Bold total row
            totalRow = numDataRows + 1;
            totalRange = Sheet.Range(sprintf('A%d:C%d', totalRow, totalRow));
            totalRange.Font.Bold = true;
            totalRange.Interior.ColorIndex = 6; % Yellow

            % Add borders
            dataRange = Sheet.Range(sprintf('A1:C%d', totalRow));
            dataRange.Borders.LineStyle = 1;

            Workbook.Save;
            Workbook.Close;
            Excel.Quit;
            delete(Excel);
        end
    catch
        % Formatting failed but file is created
    end
end
