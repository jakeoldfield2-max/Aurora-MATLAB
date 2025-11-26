function componentData = extractOccurrenceData(modelName)
% extractOccurrenceData - Extract occurrence number and properties from Systems Composer model
%
% Syntax: componentData = extractOccurrenceData(modelName)
%
% Inputs:
%    modelName - Name of the Systems Composer model (without .slx extension)
%
% Outputs:
%    componentData - Structure array containing:
%                    .OccurrenceNumber - The occurrence number
%                    .ComponentNames - Cell array of component names with this occurrence
%                    .PartNumber - Part number (if exists)
%                    .Properties - Cell array of structures containing all properties
%
% Example:
%    data = extractOccurrenceData('NRC_Template');

    % Note: The profile uses "OccuranceNumber" (typo) not "OccurrenceNumber"
    % This script handles both spellings for compatibility

    % Load the architecture model
    try
        model = systemcomposer.loadModel(modelName);
    catch ME
        error('Failed to load model: %s\nError: %s', modelName, ME.message);
    end

    % Get root architecture
    rootArch = model.Architecture;

    % Initialize data collection
    occurrenceMap = containers.Map('KeyType', 'char', 'ValueType', 'any');

    % Recursively process all components
    processComponents(rootArch, occurrenceMap);

    % Convert map to structure array for easier Excel export
    componentData = convertMapToStructArray(occurrenceMap);

end

function processComponents(arch, occurrenceMap)
% Recursively process all components in the architecture
% This function dynamically discovers all stereotype properties - not tied to specific profile names

    % Get all components at this level
    components = arch.Components;

    for i = 1:length(components)
        comp = components(i);

        % Initialize variables
        occurrenceNum = '';
        partNumber = '';
        allProperties = struct();

        % DYNAMIC APPROACH: Get all stereotype properties for this component
        % This works regardless of profile/stereotype names
        try
            propertyPaths = comp.getStereotypeProperties();

            if ~isempty(propertyPaths)
                % Iterate through all properties on this component
                for j = 1:length(propertyPaths)
                    propPath = propertyPaths(j);

                    try
                        % Get the property value
                        val = comp.getPropertyValue(propPath);

                        % Extract just the property name (last part after final dot)
                        parts = strsplit(char(propPath), '.');
                        propName = parts{end};

                        % Check if this is an OccurrenceNumber property (case-insensitive)
                        if contains(lower(propName), 'occurrencenumber') || contains(lower(propName), 'occurancenumber')
                            if ~isempty(val)
                                occurrenceNum = convertToString(val);
                            end
                        end

                        % Check if this is a PartNumber property
                        if strcmpi(propName, 'PartNumber')
                            if ~isempty(val)
                                partNumber = convertToString(val);
                            end
                        end

                        % Store all properties with valid field names
                        fieldName = sanitizeFieldName(propName);
                        allProperties.(fieldName) = val;

                    catch
                        % Skip if can't read this property
                    end
                end
            end
        catch
            % No stereotype properties on this component
        end

        % If OccurrenceNumber exists (non-empty), add this component to the map
        if ~isempty(occurrenceNum)
            compData.Name = comp.Name;
            compData.PartNumber = partNumber;
            compData.Properties = allProperties;
            try
                compData.Path = getfullname(comp.SimulinkHandle);
            catch
                compData.Path = comp.Name;
            end

            % Add to map or append to existing entry
            if occurrenceMap.isKey(occurrenceNum)
                existing = occurrenceMap(occurrenceNum);
                existing{end+1} = compData;
                occurrenceMap(occurrenceNum) = existing;
            else
                occurrenceMap(occurrenceNum) = {compData};
            end
        end

        % Recursively process child components
        try
            childArch = comp.Architecture;
            if ~isempty(childArch) && ~isempty(childArch.Components)
                processComponents(childArch, occurrenceMap);
            end
        catch
            % Skip if component doesn't have children
        end
    end
end

function structArray = convertMapToStructArray(occurrenceMap)
% Convert the occurrence map to a structured array for Excel export

    keys = occurrenceMap.keys();
    structArray = struct('OccurrenceNumber', {}, 'ComponentNames', {}, ...
                         'PartNumber', {}, 'Properties', {});

    for i = 1:length(keys)
        occNum = keys{i};
        components = occurrenceMap(occNum);

        % Collect component names
        componentNames = cell(length(components), 1);
        partNumbers = cell(length(components), 1);
        allProps = {};

        for j = 1:length(components)
            componentNames{j} = components{j}.Name;
            partNumbers{j} = components{j}.PartNumber;
            allProps{j} = components{j}.Properties;
        end

        % Use the first part number found (they should be the same for same occurrence)
        partNum = '';
        for j = 1:length(partNumbers)
            if ~isempty(partNumbers{j})
                partNum = partNumbers{j};
                break;
            end
        end

        structArray(i).OccurrenceNumber = occNum;
        structArray(i).ComponentNames = componentNames;
        structArray(i).PartNumber = partNum;
        structArray(i).Properties = allProps;
    end

    % Sort by occurrence number
    if ~isempty(structArray)
        [~, idx] = sort({structArray.OccurrenceNumber});
        structArray = structArray(idx);
    end
end

function str = convertToString(value)
% Convert various value types to string
    if ischar(value) || isstring(value)
        str = char(value);
    elseif isnumeric(value)
        str = num2str(value);
    else
        try
            str = char(value);
        catch
            str = '';
        end
    end
end

function fieldName = sanitizeFieldName(propName)
% Sanitize property name to be a valid MATLAB field name
    fieldName = propName;
    % Replace invalid characters with underscores
    fieldName = regexprep(fieldName, '[^a-zA-Z0-9_]', '_');
    % Ensure it starts with a letter
    if ~isempty(fieldName) && ~isletter(fieldName(1))
        fieldName = ['prop_' fieldName];
    end
    % Ensure it's not empty
    if isempty(fieldName)
        fieldName = 'property';
    end
end
