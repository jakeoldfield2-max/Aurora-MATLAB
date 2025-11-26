% Aurora-MATLAB Project Startup Script
% This script adds all necessary folders to the MATLAB path for System Composer profiles
%
% Run this script when opening the project, or add this folder to your MATLAB path

% Get the directory where this startup script is located
projectRoot = fileparts(mfilename('fullpath'));

% Add project folders to the path
addpath(projectRoot);
addpath(fullfile(projectRoot, 'ModelData'));
addpath(fullfile(projectRoot, 'Backend'));
addpath(fullfile(projectRoot, 'Tools'));
addpath(fullfile(projectRoot, 'PracticeFiles'));

% Display confirmation
fprintf('Aurora-MATLAB project paths added successfully.\n');
fprintf('Project root: %s\n', projectRoot);

% List available profiles
fprintf('\nAvailable System Composer Profiles:\n');
if exist(fullfile(projectRoot, 'ModelData', 'ComponentProfile.xml'), 'file')
    fprintf('  - ComponentProfile (ModelData/ComponentProfile.xml)\n');
end
if exist(fullfile(projectRoot, 'Backend', 'Components.xml'), 'file')
    fprintf('  - Components (Backend/Components.xml)\n');
end

fprintf('\nTo synchronize profiles with the model, run:\n');
fprintf('  model = systemcomposer.loadModel(''NRC_Template'');\n');
fprintf('  syncProfileToModel(model);\n');
