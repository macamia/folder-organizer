function folder_organizer_matlab_scan
% 静的な依存関係を調べ、アプリへJSONで返す。
runFolder = getenv('FOLDER_ORGANIZER_RUN');
output = fullfile(runFolder, 'result.json');
config = jsondecode(fileread(fullfile(runFolder, 'request.json')));
restoredefaultpath
cd(config.workdir)

% private、名前空間、クラスは通常のパスの下に置く。
entryFolder = fileparts(config.source);
while true
    [parent, name] = fileparts(entryFolder);
    if isempty(name) || ~(strcmpi(name, 'private') || startsWith(name, '+') || startsWith(name, '@'))
        break
    end
    entryFolder = parent;
end
folders = {entryFolder};
if ~isempty(config.searchDirs)
    folders = [cellstr(string(config.searchDirs(:))).', folders];
end
addpath(folders{:}, '-begin');

queue = {config.source};
nodes = cell(0, 1);
notices = cell(0, 1);
index = 1;
while index <= numel(queue)
    file = queue{index};
    direct = {};
    if endsWith(file, '.m', 'IgnoreCase', true)
        try
            lastwarn('');
            direct = matlab.codetools.requiredFilesAndProducts(file, 'toponly');
            [message, ~] = lastwarn;
            if ~isempty(message)
                notices{end + 1, 1} = [file ' | ' message]; %#ok<AGROW>
            end
            sourceText = regexprep(fileread(file), '(?m)^\s*%[^\r\n]*', '');
            variableTarget = '\<(feval|str2func|run|load|fopen|fileread|readtable|readmatrix)\s*\(\s*\w';
            dynamicPath = '\<(eval|evalin|addpath|rmpath|cd)\s*\(';
            if ~isempty(regexp(sourceText, [variableTarget '|' dynamicPath], 'once'))
                notices{end + 1, 1} = [file ' | 実行時に決まる参照が含まれる可能性があります。']; %#ok<AGROW>
            end
            if ispc
                direct(strcmpi(direct, file)) = [];
            else
                direct(strcmp(direct, file)) = [];
            end
            direct = unique(direct, 'stable');
        catch exception
            direct = {};
            notices{end + 1, 1} = [file ' | ' exception.message]; %#ok<AGROW>
        end
    elseif endsWith(file, {'.mlx', '.p', '.slx'}, 'IgnoreCase', true) || contains(file, '.mex')
        notices{end + 1, 1} = [file ' | この形式の内部参照は解析対象外です。']; %#ok<AGROW>
    end
    for child = reshape(direct, 1, [])
        if ispc
            queued = any(strcmpi(queue, child{1}));
        else
            queued = any(strcmp(queue, child{1}));
        end
        if ~queued
            queue{end + 1} = child{1}; %#ok<AGROW>
        end
    end
    nodes{end + 1, 1} = struct('file', file, 'dependencies', {direct}); %#ok<AGROW>
    index = index + 1;
end

report = struct('nodes', {nodes}, 'notices', {notices});
[fid, message] = fopen(output, 'w', 'n', 'UTF-8');
if fid == -1
    error('FolderOrganizer:WriteFailed', '%s', message);
end
cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
fprintf(fid, '%s', jsonencode(report));
end
