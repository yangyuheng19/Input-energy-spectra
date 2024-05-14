% Get current directory
currentDir = pwd;

% Construct file path
filePath = fullfile(currentDir, 'Fit.xlsx');

% Read Excel file data
[num, raw] = xlsread(filePath);

% Extract the first column as x and the second column as y
x = num(2:60, 1);
y = num(2:60, 2);
x1 = num(:, 1);
y1 = num(:, 2);

filePath = fullfile(currentDir, 'Fit.mat');
load(filePath);
% Set up fittype and options.
ft = fittype( 'smoothingspline' );
opts = fitoptions( 'Method', 'SmoothingSpline' );
opts.SmoothingParam = 0.9986456327510447;

% Fit model to data.
[fitresult, gof] = fit( x, y, ft, opts );
yfit =  fitresult( x );
yfit1 = [num(1:1, 2);  yfit ];
 h = plot(  x1, yfit1 );
 hold on
 h1 =  plot( x1, y1 );
 hold off
% Construct file path
filePath = fullfile(currentDir, 'yfit.xlsx');
% Add header
header = {'yfit'};

% Merge data and table headers
dataWithHeader = [header; num2cell(yfit1)];

% Write to Excel file
xlswrite(filePath, dataWithHeader, 'Sheet1'); % Write the data to column A of 'Sheet1'