function importRequirements(modelName, excelFile)
%IMPORTREQUIREMENTS Import requirements from Excel into Simulink Requirements
%
%   importRequirements() - Uses defaults
%   importRequirements(modelName) - Specify model name
%   importRequirements(modelName, excelFile) - Specify both
%
%   Imports requirements from Excel with columns:
%   A: Requirement ID
%   B: Requirement Title
%   C: Main Text (Description)
%   D: Rationale
%   E: Additional Notes
%   F: Verification Method

    % Default arguments
    if nargin < 1 || isempty(modelName)
        modelName = 'NRC_Template';
    end

    if nargin < 2 || isempty(excelFile)
        scriptPath = fileparts(mfilename('fullpath'));
        excelFile = fullfile(scriptPath, 'Aurora Requirements.xlsx');
    end

    fprintf('=== Requirements Import ===\n\n');

    % Read Excel file - read as raw data without treating first row as headers
    fprintf('Reading Excel file: %s\n', excelFile);
    T = readtable(excelFile, 'Sheet', 'NRC REQ', 'ReadVariableNames', false);

    % Check if first row is a header row and skip it
    firstCell = T{1, 1};
    if iscell(firstCell), firstCell = firstCell{1}; end
    if contains(string(firstCell), 'Requirement', 'IgnoreCase', true) || ...
       contains(string(firstCell), 'Competition', 'IgnoreCase', true)
        T(1, :) = [];  % Remove header row
        fprintf('Skipped header row\n');
    end

    fprintf('Found %d requirements\n\n', height(T));

    % Create output path for requirement set
    scriptPath = fileparts(mfilename('fullpath'));
    reqSetPath = fullfile(scriptPath, 'Aurora_Requirements.slreqx');

    % Load existing or create new requirement set
    if exist(reqSetPath, 'file')
        fprintf('Loading existing requirement set: %s\n', reqSetPath);
        rs = slreq.load(reqSetPath);
    else
        fprintf('Creating new requirement set: %s\n', reqSetPath);
        rs = slreq.new(reqSetPath);
    end

    % Import each requirement
    fprintf('Importing requirements...\n');

    createdCount = 0;
    updatedCount = 0;

    for i = 1:height(T)
        % Extract values from each column
        reqID = getTableValue(T, i, 1);
        reqTitle = getTableValue(T, i, 2);
        mainText = getTableValue(T, i, 3);
        rationale = getTableValue(T, i, 4);
        additionalNotes = getTableValue(T, i, 5);
        verificationMethod = getTableValue(T, i, 6);

        % Skip empty rows
        if isempty(reqID) || strcmp(reqID, '')
            continue;
        end

        % Skip if this looks like a header row
        if contains(reqID, 'Requirement', 'IgnoreCase', true) || ...
           contains(reqID, 'Competition', 'IgnoreCase', true)
            continue;
        end

        try
            % Check if requirement with this ID already exists
            existingReqs = find(rs, 'Type', 'Requirement', 'Id', reqID);

            if ~isempty(existingReqs)
                % Update existing requirement
                req = existingReqs(1);
                action = 'Updated';
                updatedCount = updatedCount + 1;
            else
                % Create new requirement
                req = add(rs, 'Summary', reqTitle);
                req.Id = reqID;
                action = 'Created';
                createdCount = createdCount + 1;
            end

            % Update all properties
            req.Summary = reqTitle;

            % Set description (Main Text)
            if ~isempty(mainText)
                req.Description = mainText;
            end

            % Set rationale
            if ~isempty(rationale)
                req.Rationale = rationale;
            end

            % Add verification method and additional notes to Keywords
            keywords = {};
            if ~isempty(verificationMethod)
                keywords{end+1} = ['Verification: ' verificationMethod];
            end
            if ~isempty(additionalNotes) && ~strcmp(additionalNotes, 'N/A')
                keywords{end+1} = ['Notes: ' additionalNotes];
            end
            if ~isempty(keywords)
                req.Keywords = keywords;
            end

            fprintf('  %s: %s - %s\n', action, reqID, reqTitle);

        catch ME
            fprintf('  ERROR importing %s: %s\n', reqID, ME.message);
        end
    end

    % Save the requirement set
    save(rs);

    fprintf('\n=== Import Complete ===\n');
    fprintf('Requirement set saved to: %s\n', reqSetPath);
    fprintf('  Created: %d\n', createdCount);
    fprintf('  Updated: %d\n', updatedCount);
    fprintf('  Total:   %d\n', createdCount + updatedCount);

    % Optionally link to model
    fprintf('\nTo link this requirement set to your model, run:\n');
    fprintf('  slreq.load(''Aurora_Requirements'')\n');

end

function val = getTableValue(T, row, col)
%GETTABLEVALUE Safely extract value from table cell
    try
        val = T{row, col};

        % Handle cell arrays
        if iscell(val)
            if isempty(val)
                val = '';
                return;
            end
            val = val{1};
        end

        % Handle missing values
        if any(ismissing(val), 'all')
            val = '';
            return;
        end

        % Handle NaN
        if isnumeric(val) && any(isnan(val), 'all')
            val = '';
            return;
        end

        % Convert to string
        if ~ischar(val) && ~isstring(val)
            val = char(string(val));
        end

        val = strtrim(char(val));
    catch ME
        fprintf('ERROR in getTableValue(row=%d, col=%d): %s\n', row, col, ME.message);
        val = '';
    end
end
