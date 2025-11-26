function modelAnalysis(modelName)
%MODELANALYSIS Run model analysis breakdowns with a selection dialog
%
%   modelAnalysis() - Uses default model 'NRC_Template'
%   modelAnalysis(modelName) - Uses specified model
%
%   Opens a dialog with checkboxes to select which breakdowns to run:
%   - Cost Breakdown
%   - Mass Breakdown
%   - Air Resistance Breakdown
%   - All Properties Export

    if nargin < 1
        modelName = 'NRC_Template';
    end

    % Add backend path
    scriptPath = fileparts(mfilename('fullpath'));
    backendPath = fullfile(scriptPath, '..', 'Backend', 'exportOccurrencePropertyBackend');
    addpath(backendPath);

    % Create the UI figure
    fig = uifigure('Name', 'Model Analysis', ...
                   'Position', [100 100 500 400], ...
                   'Resize', 'off');

    % Center the figure on screen
    movegui(fig, 'center');

    % Title label
    uilabel(fig, 'Position', [20 350 460 30], ...
            'Text', sprintf('Model Analysis - %s', modelName), ...
            'FontSize', 16, ...
            'FontWeight', 'bold', ...
            'HorizontalAlignment', 'center');

    % Instructions
    uilabel(fig, 'Position', [20 320 460 20], ...
            'Text', 'Select the analyses you want to run:', ...
            'FontSize', 11, ...
            'HorizontalAlignment', 'center');

    % Define analyses with descriptions
    analyses = {
        'Cost Breakdown', 'Calculates total rocket cost from all component costs. Updates CostBreakdown.xlsx';
        'Mass Breakdown', 'Calculates total rocket mass from all component masses. Updates MassBreakdown.xlsx';
        'Air Resistance Breakdown', 'Calculates total air resistance from all components. Updates AirResistanceBreakdown.xlsx';
        'All Properties Export', 'Exports all stereotype properties for all components. Updates OccurrenceProperties.xlsx'
    };

    % Create checkboxes with descriptions
    checkboxes = gobjects(size(analyses, 1), 1);
    yPos = 270;

    for i = 1:size(analyses, 1)
        % Checkbox
        checkboxes(i) = uicheckbox(fig, ...
            'Position', [30 yPos 200 22], ...
            'Text', analyses{i, 1}, ...
            'FontSize', 12, ...
            'FontWeight', 'bold', ...
            'Value', false);

        % Description label
        uilabel(fig, 'Position', [50 yPos-20 420 18], ...
                'Text', analyses{i, 2}, ...
                'FontSize', 10, ...
                'FontColor', [0.4 0.4 0.4]);

        yPos = yPos - 55;
    end

    % Select All checkbox
    selectAllCb = uicheckbox(fig, ...
        'Position', [30 yPos-10 150 22], ...
        'Text', 'Select All', ...
        'FontSize', 11, ...
        'Value', false, ...
        'ValueChangedFcn', @(src, ~) selectAllChanged(src, checkboxes));

    % Run button
    uibutton(fig, 'Position', [150 30 100 35], ...
             'Text', 'Run', ...
             'FontSize', 12, ...
             'FontWeight', 'bold', ...
             'BackgroundColor', [0.3 0.7 0.3], ...
             'FontColor', 'white', ...
             'ButtonPushedFcn', @(~, ~) runAnalyses(fig, checkboxes, modelName, scriptPath));

    % Cancel button
    uibutton(fig, 'Position', [260 30 100 35], ...
             'Text', 'Cancel', ...
             'FontSize', 12, ...
             'ButtonPushedFcn', @(~, ~) close(fig));

    % Wait for the figure to close
    uiwait(fig);
end

function selectAllChanged(src, checkboxes)
    % Toggle all checkboxes based on Select All state
    for i = 1:length(checkboxes)
        checkboxes(i).Value = src.Value;
    end
end

function runAnalyses(fig, checkboxes, modelName, scriptPath)
    % Get selected analyses
    selected = [];
    for i = 1:length(checkboxes)
        if checkboxes(i).Value
            selected(end+1) = i; %#ok<AGROW>
        end
    end

    if isempty(selected)
        uialert(fig, 'Please select at least one analysis to run.', 'No Selection');
        return;
    end

    % Close the dialog
    close(fig);

    % Ensure model is loaded
    try
        if ~bdIsLoaded(modelName)
            fprintf('Loading model: %s\n', modelName);
            load_system(modelName);
        end
    catch
        % Model might not be on path, try to find and load it
        scriptPath = fileparts(mfilename('fullpath'));
        modelPath = fullfile(scriptPath, '..', [modelName '.slx']);
        if exist(modelPath, 'file')
            fprintf('Loading model from: %s\n', modelPath);
            load_system(modelPath);
        else
            error('Could not find model: %s', modelName);
        end
    end

    % Run selected analyses
    fprintf('\n========================================\n');
    fprintf('       MODEL ANALYSIS REPORT\n');
    fprintf('       Model: %s\n', modelName);
    fprintf('========================================\n\n');

    for i = 1:length(selected)
        idx = selected(i);

        switch idx
            case 1  % Cost Breakdown
                fprintf('----------------------------------------\n');
                costBreakdown(modelName);
                fprintf('\n');

            case 2  % Mass Breakdown
                fprintf('----------------------------------------\n');
                massBreakdown(modelName);
                fprintf('\n');

            case 3  % Air Resistance Breakdown
                fprintf('----------------------------------------\n');
                airResistanceBreakdown(modelName);
                fprintf('\n');

            case 4  % All Properties Export
                fprintf('----------------------------------------\n');
                fprintf('=== All Properties Export ===\n\n');
                exportOccurrenceProperties(modelName);
                fprintf('\n');
        end
    end

    fprintf('========================================\n');
    fprintf('       ANALYSIS COMPLETE\n');
    fprintf('========================================\n');
    fprintf('\nFiles saved to: %s\n', scriptPath);

    % Show completion message
    msgbox(sprintf('Analysis complete!\n\nFiles saved to Tools folder.'), ...
           'Model Analysis Complete', 'help');
end
