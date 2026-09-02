clearvars -except SENS_*; clc; close all; rng(42);
%%
%1 Data Loading

DATE_START = '2021-06-01';
DATE_END   = '2022-06-30';

STATES    = {'ACT','NSW','NT','QLD','SA','TAS','VIC','WA'};
LOC_KEYS  = {'AU_ACT','AU_NSW','AU_NT','AU_QLD','AU_SA','AU_TAS','AU_VIC','AU_WA'};
NC        = 8;

search_dirs = {pwd};
script_file_dir = fileparts(mfilename('fullpath'));
if ~isempty(script_file_dir) && ~strcmp(script_file_dir, pwd)
    search_dirs{end+1} = script_file_dir;
end

FILE_NAMES = struct( ...
    'epi',  'Epidemiology Data.xlsx', ...
    'vacc', 'COVID-19 Vaccination by State.xlsx', ...
    'mob',  'Mobility Data.xlsx',     ...
    'demo', 'Demographics Data.xlsx');

fields     = fieldnames(FILE_NAMES);
file_paths = struct();

for fi = 1:numel(fields)
    fld   = fields{fi};
    fname = FILE_NAMES.(fld);
    found = false;
    for sd = 1:numel(search_dirs)
        candidate = fullfile(search_dirs{sd}, fname);
        if isfile(candidate)
            file_paths.(fld) = candidate;
            found = true;
            break;
        end
    end
    if ~found
        fprintf('[FILE NOT FOUND] %s\n  Please browse to select it...\n', fname);
        [fn, fp] = uigetfile('*.xlsx', sprintf('Select %s', fname));
        if isequal(fn, 0)
            error('File "%s" is required but was not selected. Aborting.', fname);
        end
        file_paths.(fld) = fullfile(fp, fn);
        if ~any(strcmp(search_dirs, fp))
            search_dirs{end+1} = fp;
        end
    end
    fprintf('  %-5s -> %s\n', upper(fld), file_paths.(fld));
end

% Population — ACT NSW NT QLD SA TAS VIC WA
% States sum to 25,683,231
Npop   = [453558;  8093815; 249200;  5217653; 1803192; 567909;  6548040; 2749864];
Npop16 = [364811;  6504442; 193211;  4161009; 1471062; 465436;  5279175; 2185967];
Nkids  = Npop - Npop16;

fprintf('\nLoading epidemiology data...\n');
epi_cols = {'date','location_key','new_confirmed','new_deceased', ...
            'new_recovered','new_tested','cum_confirmed','cum_deceased', ...
            'cum_recovered','cum_tested'};
epi_tbl  = load_covid_xlsx(file_paths.epi, epi_cols);

fprintf('Loading mobility data...\n');
mob_cols = {'date','location_key','mob_retail','mob_grocery','mob_parks', ...
            'mob_transit','mob_work','mob_reside'};
mob_tbl  = load_covid_xlsx(file_paths.mob, mob_cols);

d0     = datetime(DATE_START);
d1     = datetime(DATE_END);
dates  = d0 : d1;
Ndays  = numel(dates);


plot_xl = [datetime('2021-07-01') d1];

fprintf('  %d days from %s to %s across %d patches.\n', ...
        Ndays, DATE_START, DATE_END, NC);

new_cases  = zeros(Ndays, NC);
new_deaths = zeros(Ndays, NC);
mob_transit= zeros(Ndays, NC);
mob_work   = zeros(Ndays, NC);
mob_retail = zeros(Ndays, NC);
mob_reside = zeros(Ndays, NC);
mob_have   = false(Ndays, NC);   % true where real mobility data exists

day_num = @(dt) days(dt - d0) + 1;

for k = 1:NC
    lk = string(LOC_KEYS{k});

    sub  = epi_tbl(epi_tbl.location_key == lk, :);
    didx = day_num(sub.date);
    mask = didx >= 1 & didx <= Ndays;
    sub  = sub(mask,:);  didx = didx(mask);
    new_cases( didx, k) = max(sub.new_confirmed, 0);
    new_deaths(didx, k) = max(sub.new_deceased,  0);

    subm  = mob_tbl(mob_tbl.location_key == lk, :);
    didxm = day_num(subm.date);
    maskm = didxm >= 1 & didxm <= Ndays;
    subm  = subm(maskm,:);  didxm = didxm(maskm);
    fill0 = @(x) x .* double(~isnan(x));
    mob_transit(didxm,k) = fill0(subm.mob_transit);
    mob_work(   didxm,k) = fill0(subm.mob_work);
    mob_retail( didxm,k) = fill0(subm.mob_retail);
    mob_reside( didxm,k) = fill0(subm.mob_reside);
    mob_have(didxm, k) = ~isnan(subm.mob_work) & ~isnan(subm.mob_retail) ...
                         & ~isnan(subm.mob_transit);
end

new_cases_sm  = movmean(new_cases,  7, 1);
new_deaths_sm = movmean(new_deaths, 7, 1);
mob_transit   = movmean(mob_transit,7, 1);
mob_work      = movmean(mob_work,   7, 1);

nat_cases_obs  = zeros(Ndays, 1);
nat_deaths_obs = zeros(Ndays, 1);
sub_au  = epi_tbl(epi_tbl.location_key == "AU", :);
didx_au = day_num(sub_au.date);
mask_au = didx_au >= 1 & didx_au <= Ndays;
sub_au  = sub_au(mask_au, :);  didx_au = didx_au(mask_au);
nat_cases_obs( didx_au) = max(sub_au.new_confirmed, 0);
nat_deaths_obs(didx_au) = max(sub_au.new_deceased,  0);
nat_cases_obs_sm  = movmean(nat_cases_obs,  7);
nat_deaths_obs_sm = movmean(nat_deaths_obs, 7);

fprintf('Loading per-state vaccination data...\n');

VACC_RAMP_START = datetime('2021-02-22');
vacc_obs_v1 = zeros(Ndays, NC);
vacc_obs_v2 = zeros(Ndays, NC);
vacc_obs_v3 = zeros(Ndays, NC);

for k = 1:NC
    opts_v               = spreadsheetImportOptions('NumVariables', 4);
    opts_v.Sheet         = STATES{k};
    opts_v.VariableNames = {'Date','Dose1cum','Dose2cum','Dose3cum'};
    opts_v.DataRange     = 'A3';
    opts_v.VariableTypes = {'char','double','double','char'};
    tbl_v = readtable(file_paths.vacc, opts_v);

    nonempty = ~cellfun(@isempty, tbl_v{:,1});
    tbl_v    = tbl_v(nonempty, :);

    dates_v = datetime(string(tbl_v{:,1}), 'InputFormat', 'd MMM yyyy');

    v1_cum = tbl_v{:,2};
    v2_cum = tbl_v{:,3};
    v3_str = tbl_v{:,4};
    v3_cum = zeros(size(v3_str));
    for i = 1:numel(v3_str)
        val = str2double(v3_str{i});
        if ~isnan(val), v3_cum(i) = val; end
    end

    v1_frac = min(v1_cum / Npop16(k), 1);
    v2_frac = min(v2_cum / Npop16(k), 1);
    v3_frac = min(v3_cum / Npop16(k), 1);

    didx_v = days(dates_v - d0) + 1;
    valid  = didx_v >= 1 & didx_v <= Ndays;
    didx_v = didx_v(valid);
    v1_frac = v1_frac(valid);
    v2_frac = v2_frac(valid);
    v3_frac = v3_frac(valid);

    if numel(didx_v) >= 2
        xi = (1:Ndays)';
        vacc_obs_v1(:,k) = interp1(didx_v, v1_frac, xi, 'linear', NaN);
        vacc_obs_v2(:,k) = interp1(didx_v, v2_frac, xi, 'linear', NaN);
        vacc_obs_v3(:,k) = interp1(didx_v, v3_frac, xi, 'linear', NaN);
    end
    vacc_obs_v1(isnan(vacc_obs_v1(:,k)), k) = 0;
    vacc_obs_v2(isnan(vacc_obs_v2(:,k)), k) = 0;
    vacc_obs_v3(isnan(vacc_obs_v3(:,k)), k) = 0;

    if numel(didx_v) >= 1 && didx_v(1) > 1
        first_day  = didx_v(1);
        first_date = d0 + (first_day - 1);
        ramp_len   = max(days(first_date - VACC_RAMP_START), 1);
        for td = 1:first_day - 1
            frac = max(0, min(1, days(d0 + (td - 1) - VACC_RAMP_START) / ramp_len));
            vacc_obs_v1(td, k) = frac * vacc_obs_v1(first_day, k);
            vacc_obs_v2(td, k) = frac * vacc_obs_v2(first_day, k);
            vacc_obs_v3(td, k) = frac * vacc_obs_v3(first_day, k);
        end
    end
end
fprintf('  Per-state V1/V2/V3 observed coverage loaded.\n');

nu_dyn = zeros(3, Ndays, NC);
for k = 1:NC
    for t = 1:Ndays-1
        dv1 = max(vacc_obs_v1(t+1,k) - vacc_obs_v1(t,k), 0) * Npop16(k);
        dv2 = max(vacc_obs_v2(t+1,k) - vacc_obs_v2(t,k), 0) * Npop16(k);
        dv3 = max(vacc_obs_v3(t+1,k) - vacc_obs_v3(t,k), 0) * Npop16(k);
        % All three pools on the 16+ basis
        pool1 = max(Npop16(k)* (1 - vacc_obs_v1(t,k)),                   Npop16(k)*0.01);
        pool2 = max(Npop16(k)* (vacc_obs_v1(t,k) - vacc_obs_v2(t,k)),    10);
        pool3 = max(Npop16(k)* (vacc_obs_v2(t,k) - vacc_obs_v3(t,k)),    10);
        nu_dyn(1,t,k) = dv1 / pool1;
        nu_dyn(2,t,k) = dv2 / pool2;
        nu_dyn(3,t,k) = dv3 / pool3;
    end
    nu_dyn(:,Ndays,k) = nu_dyn(:,Ndays-1,k);
end
fprintf('  Time-varying vaccination rates (nu_dyn) derived from observed data.\n');

%%
%2 Mobility Matrix

city_lat = [-35.28, -33.87, -12.46, -27.47, -34.93, -42.88, -37.81, -31.95]; 
city_lon = [149.13, 151.21, 130.84, 153.03, 138.60, 147.33, 144.96, 115.86]; %Locations
 
dist_km = zeros(NC, NC);
for i = 1:NC
    for j = 1:NC
        if i ~= j
            dlat = deg2rad(city_lat(j) - city_lat(i));
            dlon = deg2rad(city_lon(j) - city_lon(i));
            a = sin(dlat/2)^2 + cos(deg2rad(city_lat(i))) * ...
                cos(deg2rad(city_lat(j))) * sin(dlon/2)^2;
            dist_km(i,j) = 2 * 6371 * asin(sqrt(a));
        end
    end
end
dist_km(dist_km == 0) = Inf;

phi0 = zeros(NC, NC);
for i = 1:NC
    for j = 1:NC
        if i ~= j
            phi0(i,j) = sqrt(Npop(i) * Npop(j)) / dist_km(i,j)^2;
        end
    end
end

DAILY_TRAVEL_FRAC = 0.005;
for j = 1:NC
    col_sum = sum(phi0(:,j));
    if col_sum > 0
        phi0(:,j) = phi0(:,j) * (DAILY_TRAVEL_FRAC / col_sum);
    end
end

% Inter-state border closures. TAS and WA held hard borders through 2021 (TAS
% had 5 Delta cases and 0 Delta deaths, so the gravity import must not seed
% them). border_open_day(k) is the first day state k joins inter-state travel;
% before it, row AND column k of the travel matrix are zeroed (full isolation:
% no imports into k, no exports from k). All other states are open from t=1.
border_open_day = ones(NC,1);                                  % open by default
border_open_day(strcmp(STATES,'TAS')) = days(datetime('2021-12-15') - d0) + 1;
border_open_day(strcmp(STATES,'WA'))  = days(datetime('2022-03-03') - d0) + 1;
for k = 1:NC
    if border_open_day(k) > 1
        fprintf('  Border: %s closed until day %d (%s).\n', STATES{k}, ...
                border_open_day(k), datestr(d0 + border_open_day(k) - 1, 'yyyy-mm-dd'));
    end
end

PHI = zeros(NC, NC, Ndays);
for dd = 1:Ndays
    Pd = phi0;
    for i = 1:NC
        for j = 1:NC
            if i ~= j
                si = max(1 + mob_transit(dd,i)/100, 0.05);
                sj = max(1 + mob_transit(dd,j)/100, 0.05);
                Pd(i,j) = phi0(i,j) * (si + sj) / 2;
            end
        end
    end
  
    for k = 1:NC
        if dd < border_open_day(k)
            Pd(k,:) = 0;   Pd(:,k) = 0;
        end
    end
    PHI(:,:,dd) = Pd;
end
fprintf('Mobility matrix PHI built [%dx%dx%d] with border closures.\n', NC, NC, Ndays);

%% 
%3 Parameters

mu      = 1/(70*365);
Lambda  = Npop * mu;

% Biological rates
alpha_cal_d = 0.26;   
alpha_cal_o = 0.34;    
alpha_cal   = alpha_cal_d;  
gamma_s_cal = 0.13;    
gamma_a_cal = 0.14;    
gamma_q_cal = 0.10;    
delta_cal   = 0.20;    

% CFR
CFR_d = 0.0065;  CFR_o = 0.0011;
mu_d_factor = (gamma_s_cal + delta_cal + mu) * (gamma_q_cal + mu) / ...
              (gamma_q_cal + mu + delta_cal);
mu_d_d = CFR_d * mu_d_factor;
mu_d_o = CFR_o * mu_d_factor;

% Symptomatic fraction
p_d = 0.638;  p_o = 0.700;

% Starting points
R0_d = 5.0;  R0_o = 8.2;
id_d = p_d/(gamma_s_cal + delta_cal + mu_d_d + mu) + (1-p_d)/(gamma_a_cal + mu);
id_o = p_o/(gamma_s_cal + delta_cal + mu_d_o + mu) + (1-p_o)/(gamma_a_cal + mu);
beta0_d_guess = R0_d / id_d;
beta0_o_guess = R0_o / id_o;

Reff_pre  = 0.75;   

fprintf('Initial guesses from R0: beta0_d=%.5f  beta0_o=%.5f\n', ...
        beta0_d_guess, beta0_o_guess);
fprintf('Pre-Delta (June) regime: effective R_pre=%.2f (beta0_pre set per state below)\n', Reff_pre);
fprintf('Fixed:  mu_d_d=%.8f  mu_d_o=%.8f\n', mu_d_d, mu_d_o);

% Vaccine efficacy initial guesses
epsilon1 = 0.30;  epsilon2 = 0.65;  epsilon3 = 0.87;
omega    = 1/180;

vacc_r = 1.5;     % ceiling 
vacc_w = 0.05;    % headroom 


vacc_C_hist = zeros(NC, 3);
for k = 1:NC
    vacc_C_hist(k,1) = max(vacc_obs_v1(:,k)) + vacc_w;
    vacc_C_hist(k,2) = max(vacc_obs_v2(:,k)) + vacc_w;
    vacc_C_hist(k,3) = max(vacc_obs_v3(:,k)) + vacc_w;
end


vacc_rate_sens = 0.6;   
vacc_x_hist    = 0.15;   

epsD_d = [0.72; 0.93; 0.96];   %Vaccine efficacy
epsD_o = [0.65; 0.85; 0.95];   

eps_R_fresh_d = 0.92;   
eps_R_floor_d = 0.75;   
eps_R_fresh_o = 0.80;   
eps_R_floor_o = 0.60;   
omega_NI      = 1/90; 

eps_R_death = 0.90; 
omega_R  = 1/950;
theta_ou = 0.20;  sigma_b = 0.08;  sigma_x = 0.12;  x_min = 0;

%Vaccine hesitancy replication
hes = struct();
hes.dt      = 1;
hes.x_min   = x_min;     
hes.sigma_x = sigma_x;   
hes.kappa   = 0.3;       
hes.rho     = 1/30;     
hes.x_ref   = 0.15;                         
hes.c_v0    = 0.20;      
hes.c_v_min = 0.00;     
hes.wI0     = 0.50;     
hes.wC      = 0.10;     
hes.cov_ref = 0.50;      
hes.b0      = hes.c_v0 + hes.wC*hes.cov_ref;  
hes.g_c     = 1.0;     
hes.g_w     = 0.5;      
hes.Kc      = 30;        
hes.Kd      = 0.20;     
hes.a_c     = 1.0;       
hes.a_d     = 3.0;      
hes.tau_c   = 28;       
hes.tau_d   = 28;       
                        
%Behavioural
hes.b_max   = 0.50;    
hes.ab_c    = 1.0;      
hes.ab_d    = 3.0;      

%% 
% 4 Sensitivity Analysis
if ~exist('SENS_VARIANT','var') || isempty(SENS_VARIANT)
    SENS_VARIANT = 'central';
end
switch SENS_VARIANT
    case 'central'              % all values at their central/literature settings
    case 'epsRfloor_low',       eps_R_floor_o = 0.50;
    case 'omegaNI_slow',        omega_NI      = 1/180;
    case 'omegaR_fast',         omega_R       = 1/270;
    case 'epsRdeath_low',       eps_R_death   = 0.65;
    case 'behav_weak',          hes.b_max     = 0.35;
    case 'behav_strong',        hes.b_max     = 0.65;
    case 'behav_caseweighted',  hes.ab_c      = 3.0;  hes.ab_d = 1.0;
    case 'behav_tau_short',     hes.tau_c     = 21;   hes.tau_d = 21;
    case 'noise_persistent',    theta_ou      = 0.05;
    otherwise
        error('Unknown SENS_VARIANT "%s". See the SENS block for valid names.', SENS_VARIANT);
end
fprintf('\n=== SENSITIVITY VARIANT: %s ===\n', SENS_VARIANT);
fprintf('  eps_R_floor_o=%.2f  omega_NI=1/%.0f  omega_R=1/%.0f  eps_R_death=%.2f  b_max=%.2f  ab_c/ab_d=%.1f/%.1f  tau=%d  theta_ou=%.2f\n\n', ...
    eps_R_floor_o, 1/omega_NI, 1/omega_R, eps_R_death, hes.b_max, hes.ab_c, hes.ab_d, hes.tau_c, theta_ou);

% Breakpoint dates 
DR_DAYS = [1, ...
           days(datetime('2021-07-01') - d0) + 1, ...
           days(datetime('2021-10-11') - d0) + 1, ...
           days(datetime('2021-12-01') - d0) + 1, ...
           Ndays+1];

% Detection rate 
detect_d0 = 0.35; 
detect_o0 = zeros(NC,1);
for k = 1:NC
    switch STATES{k}
        case 'NSW', detect_o0(k) = 0.622;
        case 'QLD', detect_o0(k) = 0.450;
        case 'VIC', detect_o0(k) = 0.603;
        case 'WA',  detect_o0(k) = 0.700;
        otherwise,  detect_o0(k) = 0.595;   % ACT/NT/SA/TAS: national-average fallback
    end
end
DETECT_BOUND_HALFWIDTH_D = 0.30;  
DETECT_BOUND_HALFWIDTH_O = 0.25;   
detect_lb_d = max(0.05, detect_d0 - DETECT_BOUND_HALFWIDTH_D);
detect_ub_d = min(0.95, detect_d0 + DETECT_BOUND_HALFWIDTH_D);
detect_lb_o = max(0.05, detect_o0 - DETECT_BOUND_HALFWIDTH_O);
detect_ub_o = min(0.95, detect_o0 + DETECT_BOUND_HALFWIDTH_O);
fprintf('Detection rate initial guesses: Delta=%.2f (un-anchored); Omicron per state:\n', detect_d0);
for k = 1:NC
    fprintf('  %3s: %.3f  (bounds %.2f-%.2f)\n', STATES{k}, detect_o0(k), detect_lb_o(k), detect_ub_o(k));
end

% NPI suppression index  [Ndays x NC]
NPI_SOURCE    = 'mobility';
NPI_POWER     = 2.0;                                 
NPI_W         = struct('work',0.35, 'retail',0.30, 'transit',0.35);
NPI_REF_QUANT = 0.95;    
NPI_MIN_COV   = 0.80;    
NPI_VALS = [0.50, 0.18, 0.60, 1.00];
NPI_DAYS = DR_DAYS;
npi_step = ones(Ndays, 1);
for ni = 1:4
    npi_step(NPI_DAYS(ni) : min(NPI_DAYS(ni+1)-1, Ndays)) = NPI_VALS(ni);
end
for ni = 2:4
    for td = 0:6
        rd = NPI_DAYS(ni) + td;
        if rd >= 1 && rd <= Ndays
            npi_step(rd) = NPI_VALS(ni-1) + (NPI_VALS(ni)-NPI_VALS(ni-1))*td/7;
        end
    end
end
npi_step = repmat(npi_step, 1, NC);

switch lower(NPI_SOURCE)
case 'steps'
    npi_scale = npi_step;
    fprintf('NPI source: hand-set step schedule [%s].\n', ...
            strjoin(arrayfun(@(v) sprintf('%.2f',v), NPI_VALS, ...
                             'UniformOutput', false), ' '));

case 'mobility'
    mob_retail_sm = movmean(mob_retail, 7, 1);
    act = NPI_W.work    .* (1 + mob_work     /100) ...
        + NPI_W.retail  .* (1 + mob_retail_sm/100) ...
        + NPI_W.transit .* (1 + mob_transit  /100);

    npi_scale = zeros(Ndays, NC);
    fprintf('NPI source: mobility-derived (power %.1f, weights %.2f/%.2f/%.2f).\n', ...
            NPI_POWER, NPI_W.work, NPI_W.retail, NPI_W.transit);
    for k = 1:NC
        cov_k = mean(mob_have(:,k));
        if cov_k < NPI_MIN_COV
            npi_scale(:,k) = npi_step(:,k);
            warning('F1771:mobCoverage', ...
                    ['Mobility coverage %.0f%% for %s (< %.0f%%): falling back ' ...
                     'to the step NPI schedule for this state.'], ...
                    100*cov_k, STATES{k}, 100*NPI_MIN_COV);
            fprintf('  %3s: cov %3.0f%%  -- step fallback\n', STATES{k}, 100*cov_k);
            continue;
        end

        % Interpolate gap days
        idx = find(mob_have(:,k));
        a_k = interp1(idx, act(idx,k), (1:Ndays)', 'linear');
        a_k(1:idx(1)-1)     = act(idx(1),   k);
        a_k(idx(end)+1:end) = act(idx(end), k);
        a_k = max(a_k, 0.05);          % activity multiplier stays positive

     
        a_srt = sort(a_k);
        ref_k = a_srt(max(1, min(Ndays, round(NPI_REF_QUANT*Ndays))));
        npi_scale(:,k) = (a_k / max(ref_k, eps)) .^ NPI_POWER;

        fprintf('  %3s: cov %3.0f%%  npi min %.2f  med %.2f  max %.2f\n', ...
                STATES{k}, 100*cov_k, min(npi_scale(:,k)), ...
                median(npi_scale(:,k)), max(npi_scale(:,k)));
    end
    npi_scale = min(max(npi_scale, 0.01), 1.5);

otherwise
    error('Unknown NPI_SOURCE: %s', NPI_SOURCE);
end

% Jensen's inequality correction for OU process on beta
jensen_corr = exp(-sigma_b^2 / (4 * theta_ou));

% Omicron start day
omicron_day = days(datetime('2021-12-01') - d0) + 1;
fprintf('Omicron start: day %d (%s)\n\n', omicron_day, ...
        datestr(d0 + omicron_day - 1, 'yyyy-mm-dd'));

%% 
%5 Initial Conditions

I0   = max(new_cases(1,:)' / detect_d0, 0);
E0   = 2.0 * I0;
Is0  = p_d * I0;
Ia0  = (1-p_d) * I0;
V10  = round(max((vacc_obs_v1(1,:) - vacc_obs_v2(1,:))', 0) .* Npop16);
V20  = round(max((vacc_obs_v2(1,:) - vacc_obs_v3(1,:))', 0) .* Npop16);
V30  = zeros(NC, 1);
R0v  = zeros(NC, 1);
Q0   = zeros(NC, 1);
S0   = max(Npop - E0 - Is0 - Ia0 - Q0 - R0v - V10 - V20 - V30, 0);
Sw0  = zeros(NC, 1);   % waned, previously-vaccinated susceptibles (none at t=1)
x0   = hes.x_ref * ones(NC, 1);   % start at the historical disposition (anchor)

%%
%6 Per State Calibration (16 calibrations: 8 states x 2 periods)
%  theta = [beta0, alpha, gamma_s, gamma_a, gamma_q, delta,
%           epsilon1, epsilon2, epsilon3, omega]   

fprintf('=== Per-State Calibration (10 params x 8 states x 2 periods) ===\n');
theta_bio_d = [alpha_cal_d; gamma_s_cal; gamma_a_cal; gamma_q_cal; delta_cal; ...
               epsilon1;    epsilon2;    epsilon3;    omega];
theta_bio_o = [alpha_cal_o; gamma_s_cal; gamma_a_cal; gamma_q_cal; delta_cal; ...
               epsilon1;    epsilon2;    epsilon3;    omega];

% Period-specific initial guesses 
theta0_d = [beta0_d_guess; theta_bio_d];
theta0_o = [beta0_o_guess; theta_bio_o];

% Bounds
UNPIN_BIO = false;
if UNPIN_BIO
    lb_bio = [0.10; 0.04; 0.04; 0.02; 0.01];
    ub_bio = [0.50; 0.30; 0.30; 0.15; 0.20];
else
    lb_bio = [NaN; gamma_s_cal; gamma_a_cal; gamma_q_cal; delta_cal];
    ub_bio = [NaN; gamma_s_cal; gamma_a_cal; gamma_q_cal; delta_cal];
end
lb_10   = [0.05; lb_bio; 0.05; 0.20; 0.40; 1/365];
ub_10   = [3.00; ub_bio; 0.60; 0.95; 0.99; 1/60];
lb_10_d = lb_10;  ub_10_d = ub_10;
lb_10_o = lb_10;  ub_10_o = ub_10;
if ~UNPIN_BIO
    lb_10_d(2) = alpha_cal_d;  ub_10_d(2) = alpha_cal_d;
    lb_10_o(2) = alpha_cal_o;  ub_10_o(2) = alpha_cal_o;
end

% Deaths weight 
w_death = 1;
w_cum = 10;

delta_start   = days(datetime('2021-07-01') - d0) + 1;
delta_end     = days(datetime('2021-11-30') - d0) + 1;
omicron_start = omicron_day;
omicron_end   = Ndays;

win_d = delta_start   : delta_end;
win_o = omicron_start : omicron_end;
Nd    = numel(win_d);
No    = numel(win_o);

theta_fit_d = zeros(NC, 11);   
theta_fit_o = zeros(NC, 11);
cal_d2_all  = cell(NC, 1);    
cal_o2_all  = cell(NC, 1);

%Identifiability cap on Delta beta0 
DELTA_SIGNAL_THRESH = 5;   
                           
DELTA_R0_CAP        = 1.5;    
delta_peak_pc = max(new_cases_sm(win_d,:), [], 1)' ./ (Npop/1e5);  % per-100k/day
beta0_d_cap   = (DELTA_R0_CAP / R0_d) * beta0_d_guess; 
lowsig_d      = delta_peak_pc < DELTA_SIGNAL_THRESH;   
fprintf('Delta beta0 identifiability cap (R0<=%.1f, beta0<=%.4f) applied to: %s\n', ...
    DELTA_R0_CAP, beta0_d_cap, strjoin(STATES(lowsig_d), ', '));
mu_d_d_vec  = zeros(NC, 1);  
mu_d_o_vec  = zeros(NC, 1);  
if license('test', 'Optimization_Toolbox')
    opts_lsq = optimoptions('lsqnonlin', ...
        'Display',                'iter',  ...
        'MaxIterations',          8000,    ...
        'MaxFunctionEvaluations', 400000,  ...
        'FunctionTolerance',      1e-9,    ...
        'OptimalityTolerance',    1e-9,    ...
        'StepTolerance',          1e-9);
    use_lsq = true;
else
    warning('Optimization Toolbox not available — using fminsearch.');
    use_lsq = false;
end

% Delta Calibration
fprintf('\n--- Delta period (Jul-Nov 2021) ---\n');
for k = 1:NC
    [Sk,Ek,Isk,Iak,V1k,V2k,V3k] = extract_ics(delta_start, k, p_d, ...
        new_cases_sm, detect_d0, vacc_obs_v1, vacc_obs_v2, vacc_obs_v3, Npop16, Npop);

    theta0_d_k = [theta0_d; detect_d0];
    lb_10_d_k  = [lb_10_d;  detect_lb_d];
    ub_10_d_k  = [ub_10_d;  detect_ub_d];

    if lowsig_d(k)
        ub_10_d_k(1) = min(ub_10_d_k(1), beta0_d_cap);
        theta0_d_k(1) = min(theta0_d_k(1), ub_10_d_k(1));
    end

    cal_d_base = struct( ...
        'Ndays', Nd,  'N', Npop(k),  'N16', Npop16(k),  'Lambda', Lambda(k),  'mu', mu, ...
        'p', p_d,  'epsD', epsD_d, ...
        'eps_R_fresh', eps_R_fresh_d,  'eps_R_floor', eps_R_floor_d,  'omega_NI', omega_NI, ...
        'eps_R_death', eps_R_death, ...
        'vacc_r', vacc_r,  'vacc_w', vacc_w,  'C_hist', vacc_C_hist(k,:), ...
        'vacc_rate_sens', vacc_rate_sens,  'vacc_x_hist', vacc_x_hist, ...
        'nu_dyn',          squeeze(nu_dyn(:, win_d, k)), ...
        'npi_scale',       npi_scale(win_d, k), ...
        'new_cases_sm',    new_cases_sm(win_d, k), ...
        'omega_R', omega_R,  'x_min', x_min,  'hes', hes,  'w_cum', w_cum, ...
        'S0', Sk,  'E0', Ek,  'Is0', Isk,  'Ia0', Iak, ...
        'V10', V1k,  'V20', V2k,  'V30', V3k,  'Q0', 0,  'R0v', 0,  'x0', hes.x_ref,  'Sw0', 0, ...
        'En0', Ek,  'Isn0', Isk,  'Ian0', Iak,  'Qn0', 0,  'Rn0', 0,  'Rf0', 0);

    % Pass 1: cases-only, global CFR-derived mu_d 
    fprintf('  Calibrating %s Delta (pass 1: cases only)...\n', STATES{k});
    cal_d1 = cal_d_base;
    cal_d1.mu_d = mu_d_d;  cal_d1.w_death = 0;  cal_d1.new_deaths_sm = zeros(Nd,1);
    cost_fn1 = @(th) calibration_residuals_1patch(th, cal_d1);
    r0 = cost_fn1(theta0_d_k);
    fprintf('    Initial residual norm = %.6f  (n=%d elements)\n', norm(r0), numel(r0));
    if use_lsq
        th_fit1 = lsqnonlin(cost_fn1, theta0_d_k, lb_10_d_k, ub_10_d_k, opts_lsq);
    else
        th_fit1 = fminsearch(@(th) sum(cost_fn1(th).^2), theta0_d_k);
    end

    % Re-derive mu_d from pass-1
    mu_d_factor_k  = (th_fit1(3)+th_fit1(6)+mu) * (th_fit1(5)+mu) / (th_fit1(5)+mu+th_fit1(6));
    mu_d_d_vec(k)  = CFR_d * mu_d_factor_k;

    % Pass 2: cases + deaths (equal weight), state-specific mu_d
    fprintf('  Calibrating %s Delta (pass 2: cases+deaths, mu_d=%.8f)...\n', STATES{k}, mu_d_d_vec(k));
    cal_d2 = cal_d_base;
    cal_d2.mu_d = mu_d_d_vec(k);  cal_d2.w_death = w_death;
    cal_d2.new_deaths_sm = new_deaths_sm(win_d, k);
    cost_fn2 = @(th) calibration_residuals_1patch(th, cal_d2);
    if use_lsq
        [th_fit, rn, ~, exitflag] = lsqnonlin(cost_fn2, th_fit1, lb_10_d_k, ub_10_d_k, opts_lsq);
    else
        th_fit = fminsearch(@(th) sum(cost_fn2(th).^2), th_fit1);
        rn = sum(cost_fn2(th_fit).^2);  exitflag = NaN;
    end
    theta_fit_d(k,:) = th_fit';
    cal_d2_all{k} = cal_d2;   % retained for the calibration-vs-simulation diagnostic
    fprintf('  %3s Delta: beta0=%.4f  alpha=%.4f  gs=%.4f  eps1=%.3f  eps2=%.3f  rn=%.4f  exit=%d\n', ...
        STATES{k}, th_fit(1), th_fit(2), th_fit(3), th_fit(7), th_fit(8), rn, exitflag);
end


fprintf('\n--- Delta forward pass for Omicron ICs ---\n');

Nfwd   = omicron_start - delta_start;
R_fwd  = zeros(NC, 1);  Rn_fwd = zeros(NC, 1);  Rf_fwd = zeros(NC, 1);
S_fwd  = zeros(NC, 1);  E_fwd  = zeros(NC, 1);  Sw_fwd = zeros(NC, 1);
Is_fwd = zeros(NC, 1);  Ia_fwd = zeros(NC, 1);
V1_fwd = zeros(NC, 1);  V2_fwd = zeros(NC, 1);  V3_fwd = zeros(NC, 1);

for k = 1:NC
    % Extract fitted params (theta = [beta0, alpha, gs, ga, gq, delta, e1, e2, e3, omega])
    beta0_k   = theta_fit_d(k,1);
    alpha_k   = theta_fit_d(k,2);  gamma_s_k = theta_fit_d(k,3);
    gamma_a_k = theta_fit_d(k,4);  gamma_q_k = theta_fit_d(k,5);
    delta_k   = theta_fit_d(k,6);  eps1_k    = theta_fit_d(k,7);
    omega_k   = theta_fit_d(k,10);
    detect_k  = theta_fit_d(k,11);   % fitted Delta detect rate for this state

    [Sk,Ek,Isk,Iak,V1k,V2k,V3k] = extract_ics(delta_start, k, p_d, ...
        new_cases_sm, detect_d0, vacc_obs_v1, vacc_obs_v2, vacc_obs_v3, Npop16, Npop);
    Qk = 0;  Rvk = 0;  xk = hes.x_ref;  Swk = 0;  
    Enk = Ek;  Isnk = Isk;  Iank = Iak;  Qnk = 0;  Rnk = 0;
    Sck = min(Nkids(k), Sk);   % under-16 susceptibles 
    Rfk = 0;   

    Mc = 1e5*max(p_d*alpha_k*Ek*detect_k,0)/Npop(k);
    Md = 1e5*max(mu_d_d_vec(k)*(Isk+Qk),0)/Npop(k);

    for ti = 1:Nfwd
        t_abs = delta_start + ti - 1;
        Nk    = max(Sk+Swk+Ek+Isk+Iak+Qk+Rvk+V1k+V2k+V3k, 1);
        lam   = beta0_k * npi_scale(t_abs, k) * behav_suppression(Mc, Md, hes) * (Isk+Iak) / Nk;
        l1    = (1 - eps1_k) * lam;

        x_disp = hes.x_ref;  
        g_up = vacc_uptake(x_disp, vacc_rate_sens, vacc_x_hist);   % demand speed multiplier
        nu1e = nu_dyn(1,t_abs,k) * g_up * vacc_headroom((V1k+V2k+V3k)/Npop16(k), x_disp, vacc_r, vacc_w, vacc_C_hist(k,1), vacc_x_hist);
        nu2e = nu_dyn(2,t_abs,k) * g_up * vacc_headroom((V2k+V3k)/Npop16(k),     x_disp, vacc_r, vacc_w, vacc_C_hist(k,2), vacc_x_hist);
        nu3e = nu_dyn(3,t_abs,k) * g_up * vacc_headroom(V3k/Npop16(k),           x_disp, vacc_r, vacc_w, vacc_C_hist(k,3), vacc_x_hist);

        Rwk = max(0, Rvk - Rnk);   % recovered who had been vaccinated before infection
        Sek = max(0, Sk - Sck);    % eligible (16+) susceptibles: first doses only
        dSc = Lambda(k) - lam*Sck - mu*Sck;   % under-16 subset of S
        dS  = Lambda(k) - lam*Sk - nu1e*Sek - mu*Sk + omega_R*Rnk;
        dSw = omega_k*V3k - lam*Swk - mu*Swk + omega_R*Rwk;   % booster + vacc-recovered waning
        % Waning natural immunity
        Ruk = max(0, Rvk - Rfk);   % waned (natural-immunity floor) recovered
        lamR_fresh = lam * (1 - eps_R_fresh_d);
        lamR_floor = lam * (1 - eps_R_floor_d);
        reinf_tot   = lamR_fresh*Rfk + lamR_floor*Ruk;
        reinf_naive = reinf_tot * (Rnk / max(Rvk,1));
        dE  = lam*Sk + lam*Swk + reinf_tot + l1*V1k - alpha_k*Ek - mu*Ek;
        dIs = p_d*alpha_k*Ek - (gamma_s_k+delta_k+mu+mu_d_d_vec(k))*Isk;
        dIa = (1-p_d)*alpha_k*Ek - (gamma_a_k+mu)*Iak;
        dQ  = delta_k*Isk - (gamma_q_k+mu+mu_d_d_vec(k))*Qk;
        dR  = gamma_s_k*Isk + gamma_a_k*Iak + gamma_q_k*Qk - mu*Rvk - omega_R*Rvk - reinf_tot;
        dRf = gamma_s_k*Isk + gamma_a_k*Iak + gamma_q_k*Qk - mu*Rfk - omega_R*Rfk - omega_NI*Rfk - lamR_fresh*Rfk;

        % Shadow chain
        dEn  = lam*Sk + reinf_naive - alpha_k*Enk - mu*Enk;
        dIsn = p_d*alpha_k*Enk - (gamma_s_k+delta_k+mu+mu_d_d_vec(k))*Isnk;
        dIan = (1-p_d)*alpha_k*Enk - (gamma_a_k+mu)*Iank;
        dQn  = delta_k*Isnk - (gamma_q_k+mu+mu_d_d_vec(k))*Qnk;
        dRn  = gamma_s_k*Isnk + gamma_a_k*Iank + gamma_q_k*Qnk - mu*Rnk - omega_R*Rnk - reinf_naive;
        dV1 = nu1e*Sek - nu2e*V1k - mu*V1k;
        dV2 = nu2e*V1k - nu3e*V2k - mu*V2k;
        dV3 = nu3e*V2k - omega_k*V3k - mu*V3k;

        Sk  = max(0, Sk+dS);   Ek  = max(0, Ek+dE);
        Swk = max(0, Swk+dSw);
        Isk = max(0, Isk+dIs); Iak = max(0, Iak+dIa);
        Qk  = max(0, Qk+dQ);   Rvk = max(0, Rvk+dR);
        V1k = max(0, V1k+dV1); V2k = max(0, V2k+dV2); V3k = max(0, V3k+dV3);
        Enk  = min(max(0, Enk+dEn),   Ek);   Isnk = min(max(0, Isnk+dIsn), Isk);
        Iank = min(max(0, Iank+dIan), Iak);  Qnk  = min(max(0, Qnk+dQn),   Qk);
        Rnk  = min(max(0, Rnk+dRn),   Rvk);
        Sck  = min(max(0, Sck+dSc),   Sk);   
        Rfk  = min(max(0, Rfk+dRf),   Rvk); 

        vacc_cov   = (V1k+V2k+V3k)/Nk;
        cases_now  = p_d*alpha_k*Ek*detect_k;
        deaths_now = mu_d_d_vec(k)*(Isk+Qk);
        [xk, Mc, Md] = hesitancy_step(xk, cases_now, deaths_now, vacc_cov, ...
                                      Npop(k), Mc, Md, 0, hes);
    end

    R_fwd(k)  = Rvk; Rn_fwd(k) = Rnk; Rf_fwd(k) = Rfk;
    S_fwd(k)  = Sk;  E_fwd(k)  = Ek;  Sw_fwd(k) = Swk;
    Is_fwd(k) = Isk; Ia_fwd(k) = Iak;
    V1_fwd(k) = V1k; V2_fwd(k) = V2k;  V3_fwd(k) = V3k;
    fprintf('  %3s: R_fwd=%.0f  S_fwd=%.0f\n', STATES{k}, Rvk, Sk);
end

% Omicron Calibration
fprintf('\n--- Omicron period (Dec 2021 - Jun 2022) ---\n');
for k = 1:NC
    I0k  = max(new_cases_sm(omicron_start, k) / detect_o0(k), 0);
    E0k  = 2.0 * I0k;
    Is0k = p_o * I0k;
    Ia0k = (1-p_o) * I0k;
    Q0k  = 0;
    Sw0k = Sw_fwd(k);
    S0k  = max(Npop(k) - E0k - Is0k - Ia0k - Q0k - R_fwd(k) - Sw0k ...
               - V1_fwd(k) - V2_fwd(k) - V3_fwd(k), 0);
    % The Omicron infection seed is re-drawn from observed cases
    naive_share = S0k / max(S0k + Sw0k + V1_fwd(k) + V2_fwd(k) + V3_fwd(k), 1);
    Rn0k = Rn_fwd(k);
    Rf0k = min(Rf_fwd(k), R_fwd(k));   % freshly-recovered natural immunity carries over
    % guess/bounds as theta element 11 for this state's calibration.
    theta0_o_k = [theta0_o; detect_o0(k)];
    lb_10_o_k  = [lb_10_o;  detect_lb_o(k)];
    ub_10_o_k  = [ub_10_o;  detect_ub_o(k)];

    cal_o_base = struct( ...
        'Ndays', No,  'N', Npop(k),  'N16', Npop16(k),  'Lambda', Lambda(k),  'mu', mu, ...
        'p', p_o,  'epsD', epsD_o, ...
        'eps_R_fresh', eps_R_fresh_o,  'eps_R_floor', eps_R_floor_o,  'omega_NI', omega_NI, ...
        'eps_R_death', eps_R_death, ...
        'vacc_r', vacc_r,  'vacc_w', vacc_w,  'C_hist', vacc_C_hist(k,:), ...
        'vacc_rate_sens', vacc_rate_sens,  'vacc_x_hist', vacc_x_hist, ...
        'nu_dyn',          squeeze(nu_dyn(:, win_o, k)), ...
        'npi_scale',       npi_scale(win_o, k), ...
        'new_cases_sm',    new_cases_sm(win_o, k), ...
        'omega_R', omega_R,  'x_min', x_min,  'hes', hes,  'w_cum', w_cum, ...
        'S0', S0k,  'E0', E0k,  'Is0', Is0k,  'Ia0', Ia0k, ...
        'V10', V1_fwd(k),  'V20', V2_fwd(k),  'V30', V3_fwd(k), ...
        'Q0', Q0k,  'R0v', R_fwd(k),  'x0', hes.x_ref,  'Sw0', Sw0k, ...
        'En0', naive_share*E0k,  'Isn0', naive_share*Is0k, ...
        'Ian0', naive_share*Ia0k,  'Qn0', naive_share*Q0k,  'Rn0', Rn0k,  'Rf0', Rf0k);

    % Pass 1: cases-only, global CFR-derived mu_d.
    fprintf('  Calibrating %s Omicron (pass 1: cases only)...\n', STATES{k});
    cal_o1 = cal_o_base;
    cal_o1.mu_d = mu_d_o;  cal_o1.w_death = 0;  cal_o1.new_deaths_sm = zeros(No,1);
    cost_fn1 = @(th) calibration_residuals_1patch(th, cal_o1);
    r0 = cost_fn1(theta0_o_k);
    fprintf('    Initial residual norm = %.6f  (n=%d elements)\n', norm(r0), numel(r0));
    if use_lsq
        th_fit1 = lsqnonlin(cost_fn1, theta0_o_k, lb_10_o_k, ub_10_o_k, opts_lsq);
    else
        th_fit1 = fminsearch(@(th) sum(cost_fn1(th).^2), theta0_o_k);
    end

    % Re-derive mu_d from pass-1
    mu_d_factor_k  = (th_fit1(3)+th_fit1(6)+mu) * (th_fit1(5)+mu) / (th_fit1(5)+mu+th_fit1(6));
    mu_d_o_vec(k)  = CFR_o * mu_d_factor_k;

    % Pass 2: cases + deaths (equal weight), state-specific mu_d
    fprintf('  Calibrating %s Omicron (pass 2: cases+deaths, mu_d=%.8f)...\n', STATES{k}, mu_d_o_vec(k));
    cal_o2 = cal_o_base;
    cal_o2.mu_d = mu_d_o_vec(k);  cal_o2.w_death = w_death;
    cal_o2.new_deaths_sm = new_deaths_sm(win_o, k);
    cost_fn2 = @(th) calibration_residuals_1patch(th, cal_o2);
    if use_lsq
        [th_fit, rn, ~, exitflag] = lsqnonlin(cost_fn2, th_fit1, lb_10_o_k, ub_10_o_k, opts_lsq);
    else
        th_fit = fminsearch(@(th) sum(cost_fn2(th).^2), th_fit1);
        rn = sum(cost_fn2(th_fit).^2);  exitflag = NaN;
    end
    theta_fit_o(k,:) = th_fit';
    cal_o2_all{k} = cal_o2;
    fprintf('  %3s Omicron: beta0=%.4f  alpha=%.4f  gs=%.4f  eps1=%.3f  eps2=%.3f  rn=%.4f  exit=%d\n', ...
        STATES{k}, th_fit(1), th_fit(2), th_fit(3), th_fit(7), th_fit(8), rn, exitflag);
end

fprintf('\n========== Per-State Fitted Parameters ==========================\n');
fprintf('%-4s  %7s %7s %7s %7s %7s %7s | %7s %7s %7s %7s\n', ...
    'State','beta0_d','alpha_d','gs_d','eps1_d','eps2_d','detR_d','beta0_o','alpha_o','gs_o','detR_o');
for k = 1:NC
    fprintf('%-4s  %7.4f %7.4f %7.4f %7.4f %7.4f %7.3f | %7.4f %7.4f %7.4f %7.3f\n', ...
        STATES{k}, theta_fit_d(k,1), theta_fit_d(k,2), theta_fit_d(k,3), ...
        theta_fit_d(k,7), theta_fit_d(k,8), theta_fit_d(k,11), ...
        theta_fit_o(k,1), theta_fit_o(k,2), theta_fit_o(k,3), theta_fit_o(k,11));
end
fprintf('=================================================================\n\n');

%% 
%7 Parameter Matrices
fprintf('Building parameter matrices...\n');

Nsteps = Ndays;
dt     = 1;

step_mat = @(vd, vo) [repmat(vd', omicron_day-1, 1); ...
                      repmat(vo', Nsteps-omicron_day+1, 1)];

% [Nsteps x NC] matrices
% 1=beta0, 2=alpha, 3=gamma_s, 4=gamma_a, 5=gamma_q, 6=delta, 7=epsilon1, 8=epsilon2, 9=epsilon3, 10=omega
beta0_mat   = step_mat(theta_fit_d(:,1),  theta_fit_o(:,1));
alpha_mat   = step_mat(theta_fit_d(:,2),  theta_fit_o(:,2));
gamma_s_mat = step_mat(theta_fit_d(:,3),  theta_fit_o(:,3));
gamma_a_mat = step_mat(theta_fit_d(:,4),  theta_fit_o(:,4));
gamma_q_mat = step_mat(theta_fit_d(:,5),  theta_fit_o(:,5));
delta_mat   = step_mat(theta_fit_d(:,6),  theta_fit_o(:,6));
eps1_mat    = step_mat(theta_fit_d(:,7),  theta_fit_o(:,7));
eps2_mat    = step_mat(theta_fit_d(:,8),  theta_fit_o(:,8));
eps3_mat    = step_mat(theta_fit_d(:,9),  theta_fit_o(:,9));
omega_mat   = step_mat(theta_fit_d(:,10), theta_fit_o(:,10));
detect_rate_mat = step_mat(theta_fit_d(:,11), theta_fit_o(:,11));   % fitted, per-state/period detection rate

% [Nsteps x 1] vectors for fixed params that switch at Omicron
step_vec = @(vd, vo) [vd*ones(omicron_day-1,1); vo*ones(Nsteps-omicron_day+1,1)];
p_vec    = step_vec(p_d,    p_o);

% [Nsteps x NC] per-state death rate
% fitted removal rates
mu_d_mat = step_mat(mu_d_d_vec, mu_d_o_vec);
epsD1_vec = step_vec(epsD_d(1), epsD_o(1));
epsD2_vec = step_vec(epsD_d(2), epsD_o(2));
epsD3_vec = step_vec(epsD_d(3), epsD_o(3));
eps_R_fresh_vec = step_vec(eps_R_fresh_d, eps_R_fresh_o); 
eps_R_floor_vec = step_vec(eps_R_floor_d, eps_R_floor_o);   

% beta0_pre is set per state
beta0_pre = zeros(1, NC);
fprintf('Pre-Delta June beta0_pre (effective R_pre=%.2f):\n', Reff_pre);
for k = 1:NC
    npi_pre_k    = mean(npi_scale(1:delta_start-1, k));
    beta0_pre(k) = Reff_pre / id_d / max(npi_pre_k, eps);
    fprintf('  %3s: mean June npi=%.2f  beta0_pre=%.5f\n', ...
            STATES{k}, npi_pre_k, beta0_pre(k));
end
beta0_mat(1:delta_start-1, :) = repmat(beta0_pre, delta_start-1, 1);

% [Nsteps x NC] beta_base with Jensen correction
beta_base_mat = beta0_mat * jensen_corr;

%% 
%8 Stochastic Simulation
Nrep = 100;
fprintf('Running stochastic multi-city simulation (%d-run ensemble)...\n', Nrep);

% Deterministic per-step matrices 
dr_mat   = detect_rate_mat;   
p_mat    = repmat(p_vec, 1, NC);

% Per-run compartment arrays
S_arr  = zeros(Nsteps,NC);  E_arr  = zeros(Nsteps,NC);
Sw_arr = zeros(Nsteps,NC);
Is_arr = zeros(Nsteps,NC);  Ia_arr = zeros(Nsteps,NC);
Q_arr  = zeros(Nsteps,NC);  R_arr  = zeros(Nsteps,NC);
V1_arr = zeros(Nsteps,NC);  V2_arr = zeros(Nsteps,NC);
V3_arr = zeros(Nsteps,NC);
x_arr  = zeros(Nsteps,NC);
% Shadow chain
En_arr = zeros(Nsteps,NC);  Isn_arr = zeros(Nsteps,NC);
Ian_arr= zeros(Nsteps,NC);  Qn_arr  = zeros(Nsteps,NC);
Rn_arr = zeros(Nsteps,NC);
Sc_arr = zeros(Nsteps,NC);  % under-16 susceptibles
Rf_arr = zeros(Nsteps,NC);  % freshly-recovered natural immunity 
xi_b   = zeros(Nsteps,NC);
beta_t = zeros(Nsteps,NC);
cumV1_arr = zeros(Nsteps,NC);
cumV2_arr = zeros(Nsteps,NC);
cumV3_arr = zeros(Nsteps,NC);
Mc_arr = zeros(Nsteps,NC);  Md_arr = zeros(Nsteps,NC);   
mu_d_eff_arr = mu_d_mat; 

% Ensemble storage [Nsteps x NC x Nrep]
NewCases_ens  = zeros(Nsteps, NC, Nrep);
NewDeaths_ens = zeros(Nsteps, NC, Nrep);

for r = 1:Nrep
    rng(42 + r);

    % Reset all per-run arrays for this replicate
    S_arr(:)=0;  E_arr(:)=0;  Sw_arr(:)=0; Is_arr(:)=0; Ia_arr(:)=0;
    Q_arr(:)=0;  R_arr(:)=0;  V1_arr(:)=0; V2_arr(:)=0; V3_arr(:)=0;
    En_arr(:)=0; Isn_arr(:)=0; Ian_arr(:)=0; Qn_arr(:)=0; Rn_arr(:)=0;
    Sc_arr(:)=0; Rf_arr(:)=0;
    x_arr(:)=0;  xi_b(:)=0;   beta_t(:)=0;
    cumV1_arr(:)=0; cumV2_arr(:)=0; cumV3_arr(:)=0;
    Mc_arr(:)=0; Md_arr(:)=0;  mu_d_eff_arr = mu_d_mat;

% NPI suppression (npi_scale) is applied
sim_start = 1;
for k = 1:NC
    [Sk,Ek,Isk,Iak,V1k,V2k,V3k] = extract_ics(sim_start, k, p_d, ...
        new_cases_sm, detect_d0, vacc_obs_v1, vacc_obs_v2, vacc_obs_v3, Npop16, Npop);
    S_arr(sim_start,k)  = Sk;   E_arr(sim_start,k)  = Ek;
    Sw_arr(sim_start,k) = 0;
    Is_arr(sim_start,k) = Isk;  Ia_arr(sim_start,k) = Iak;
    Q_arr(sim_start,k)  = 0;    R_arr(sim_start,k)  = 0;
    V1_arr(sim_start,k) = V1k;  V2_arr(sim_start,k) = V2k;
    V3_arr(sim_start,k) = V3k;  x_arr(sim_start,k)  = hes.x_ref;   % start at the anchor
    En_arr(sim_start,k) = Ek;   Isn_arr(sim_start,k) = Isk;
    Ian_arr(sim_start,k)= Iak;  Qn_arr(sim_start,k)  = 0;
    Rn_arr(sim_start,k) = 0;
    Sc_arr(sim_start,k) = min(Nkids(k), Sk);   % all under-16s start susceptible
    Rf_arr(sim_start,k) = 0;   % no recoveries yet
    cumV1_arr(sim_start,k) = V1k + V2k + V3k;
    cumV2_arr(sim_start,k) = V2k + V3k;
    cumV3_arr(sim_start,k) = V3k;
    Mc_arr(sim_start,k) = 1e5*max(p_d*alpha_mat(sim_start,k)*Ek*detect_rate_mat(sim_start,k),0)/Npop(k);
    Md_arr(sim_start,k) = 1e5*max(mu_d_d_vec(k)*Isk,0)/Npop(k);
end
xi_b(sim_start,:)   = 0;
beta_t(sim_start,:) = beta_base_mat(sim_start,:);

for t = sim_start:Nsteps-1
    % OU update
    xi_b(t+1,:)   = xi_b(t,:) - theta_ou*xi_b(t,:)*dt + sigma_b*sqrt(dt)*randn(1,NC);
    beta_t(t+1,:) = beta_base_mat(t+1,:) .* npi_scale(t+1,:) .* exp(xi_b(t+1,:));

    Phi = PHI(:,:, min(t, Ndays));

    for k = 1:NC
        S  = S_arr(t,k);   E  = E_arr(t,k);
        Sw = Sw_arr(t,k);
        Is = Is_arr(t,k);  Ia = Ia_arr(t,k);
        Q  = Q_arr(t,k);   Rv = R_arr(t,k);
        En = En_arr(t,k);  Isn = Isn_arr(t,k);  Ian = Ian_arr(t,k);
        Qn = Qn_arr(t,k);  Rn  = Rn_arr(t,k);
        Sc = Sc_arr(t,k);
        Rf = Rf_arr(t,k);
        V1 = V1_arr(t,k);  V2 = V2_arr(t,k);  V3 = V3_arr(t,k);
        xk = x_arr(t,k);
        N  = max(S+Sw+E+Is+Ia+Q+Rv+V1+V2+V3, 1);
        beta_now = beta_t(t+1, k) * behav_suppression(Mc_arr(t,k), Md_arr(t,k), hes);

        % Per-state, per-period parameters
        alpha_k   = alpha_mat(t+1, k);
        gamma_s_k = gamma_s_mat(t+1, k);
        gamma_a_k = gamma_a_mat(t+1, k);
        gamma_q_k = gamma_q_mat(t+1, k);
        delta_k   = delta_mat(t+1, k);
        eps1_k    = eps1_mat(t+1, k);
        eps2_k    = eps2_mat(t+1, k);
        eps3_k    = eps3_mat(t+1, k);
        omega_k   = omega_mat(t+1, k);
        p_k       = p_vec(t+1);
        mu_d_k    = mu_d_mat(t+1, k);

        % Mobility coupling
        mob_inflow = 0;
        for j = 1:NC
            if j ~= k
                mob_inflow = mob_inflow + Phi(k,j) * (Is_arr(t,j) + Ia_arr(t,j));
            end
        end
        lam_mob = beta_now * mob_inflow / N;

        % Force of infection
        lam     = beta_now * (Is + Ia) / N;
        lam_tot = lam + lam_mob;
        lam1    = (1 - eps1_k) * lam_tot;
        lam2    = (1 - eps2_k) * lam_tot;
        lam3    = (1 - eps3_k) * lam_tot;

        % Reinfection flow
        Ru = max(0, Rv - Rf);
        lamR_fresh = lam_tot * (1 - eps_R_fresh_vec(t+1));
        lamR_floor = lam_tot * (1 - eps_R_floor_vec(t+1));
        reinf_tot   = lamR_fresh*Rf + lamR_floor*Ru;
        reinf_naive = reinf_tot * (Rn / max(Rv,1));

        % Vaccine and reinfection weighted effective death rate
        mu_d_eff = death_rate_eff(mu_d_k, lam_tot*(S+Sw), lam1*V1, lam2*V2, lam3*V3, ...
                                  [epsD1_vec(t+1); epsD2_vec(t+1); epsD3_vec(t+1)], ...
                                  reinf_tot, eps_R_death);
        mu_d_eff_arr(t,k) = mu_d_eff;

        % Vaccine hesitancy 
        vacc_cov   = (V1 + V2 + V3) / N;
        cases_now  = p_k*alpha_k*E*detect_rate_mat(t,k);   % modelled reported cases
        deaths_now = mu_d_eff*(Is + Q);                  % modelled deaths
        dW_x       = sqrt(dt) * randn;
        [x_arr(t+1,k), Mc_arr(t+1,k), Md_arr(t+1,k)] = hesitancy_step( ...
            xk, cases_now, deaths_now, vacc_cov, N, Mc_arr(t,k), Md_arr(t,k), dW_x, hes);

        % Hesitancy-adjusted vaccination rates
  
        x_disp = hes.x_ref;
        g_up = vacc_uptake(x_disp, vacc_rate_sens, vacc_x_hist);   % demand speed multiplier
        nu1e = nu_dyn(1,t,k) * g_up * vacc_headroom((V1+V2+V3)/Npop16(k), x_disp, vacc_r, vacc_w, vacc_C_hist(k,1), vacc_x_hist);
        nu2e = nu_dyn(2,t,k) * g_up * vacc_headroom((V2+V3)/Npop16(k),    x_disp, vacc_r, vacc_w, vacc_C_hist(k,2), vacc_x_hist);
        nu3e = nu_dyn(3,t,k) * g_up * vacc_headroom(V3/Npop16(k),         x_disp, vacc_r, vacc_w, vacc_C_hist(k,3), vacc_x_hist);

        % Compartment ODEs
        Rw  = max(0, Rv - Rn);   
        Se  = max(0, S - Sc);   
        dS  = (Lambda(k) - lam_tot*S - nu1e*Se - mu*S + omega_R*Rn) * dt;
        dSc = (Lambda(k) - lam_tot*Sc - mu*Sc) * dt;
        dSw = (omega_k*V3 - lam_tot*Sw - mu*Sw + omega_R*Rw) * dt;  
        dE  = (lam_tot*S + lam_tot*Sw + reinf_tot + lam1*V1 + lam2*V2 + lam3*V3 - alpha_k*E - mu*E) * dt;
        dIs = (p_k*alpha_k*E - (gamma_s_k+delta_k+mu+mu_d_eff)*Is) * dt;
        dIa = ((1-p_k)*alpha_k*E - (gamma_a_k+mu)*Ia) * dt;
        dQ  = (delta_k*Is - (gamma_q_k+mu+mu_d_eff)*Q) * dt;
        dR  = (gamma_s_k*Is + gamma_a_k*Ia + gamma_q_k*Q - mu*Rv - omega_R*Rv - reinf_tot) * dt;
        dRf = (gamma_s_k*Is + gamma_a_k*Ia + gamma_q_k*Q - mu*Rf - omega_R*Rf - omega_NI*Rf - lamR_fresh*Rf) * dt;

        % Shadow chain
        dEn  = (lam_tot*S + reinf_naive - alpha_k*En - mu*En) * dt;
        dIsn = (p_k*alpha_k*En - (gamma_s_k+delta_k+mu+mu_d_eff)*Isn) * dt;
        dIan = ((1-p_k)*alpha_k*En - (gamma_a_k+mu)*Ian) * dt;
        dQn  = (delta_k*Isn - (gamma_q_k+mu+mu_d_eff)*Qn) * dt;
        dRn  = (gamma_s_k*Isn + gamma_a_k*Ian + gamma_q_k*Qn - mu*Rn - omega_R*Rn - reinf_naive) * dt;
        dV1 = (nu1e*Se - nu2e*V1 - mu*V1) * dt;
        dV2 = (nu2e*V1 - nu3e*V2 - mu*V2) * dt;
        dV3 = (nu3e*V2 - omega_k*V3 - mu*V3) * dt;

        S_arr(t+1,k)  = max(0, S+dS);    E_arr(t+1,k)  = max(0, E+dE);
        Sw_arr(t+1,k) = max(0, Sw+dSw);
        Is_arr(t+1,k) = max(0, Is+dIs);  Ia_arr(t+1,k) = max(0, Ia+dIa);
        Q_arr(t+1,k)  = max(0, Q+dQ);    R_arr(t+1,k)  = max(0, Rv+dR);
        V1_arr(t+1,k) = max(0, V1+dV1);  V2_arr(t+1,k) = max(0, V2+dV2);
        V3_arr(t+1,k) = max(0, V3+dV3);
        En_arr(t+1,k)  = min(max(0, En+dEn),   E_arr(t+1,k));
        Isn_arr(t+1,k) = min(max(0, Isn+dIsn), Is_arr(t+1,k));
        Ian_arr(t+1,k) = min(max(0, Ian+dIan), Ia_arr(t+1,k));
        Qn_arr(t+1,k)  = min(max(0, Qn+dQn),   Q_arr(t+1,k));
        Rn_arr(t+1,k)  = min(max(0, Rn+dRn),   R_arr(t+1,k));
        Sc_arr(t+1,k)  = min(max(0, Sc+dSc),   S_arr(t+1,k));   
        Rf_arr(t+1,k)  = min(max(0, Rf+dRf),   R_arr(t+1,k));  
        cumV1_arr(t+1,k) = cumV1_arr(t,k) + nu1e*Se*dt;
        cumV2_arr(t+1,k) = cumV2_arr(t,k) + nu2e*V1*dt;
        cumV3_arr(t+1,k) = cumV3_arr(t,k) + nu3e*V2*dt;
    end
    end

    % Boundary fix
    mu_d_eff_arr(Nsteps,:) = mu_d_eff_arr(Nsteps-1,:);

    % Store this replicate's modelled cases/deaths
    NewCases_ens(:,:,r)  = p_mat .* alpha_mat .* E_arr .* dr_mat;
    NewDeaths_ens(:,:,r) = mu_d_eff_arr .* (Is_arr + Q_arr);
end   % replicate loop

% Ensemble summaries
new_cases_model  = median(NewCases_ens, 3);
new_deaths_model = median(NewDeaths_ens, 3);
cases_q05  = quantile(NewCases_ens,  0.05, 3);
cases_q95  = quantile(NewCases_ens,  0.95, 3);
deaths_q05 = quantile(NewDeaths_ens, 0.05, 3);
deaths_q95 = quantile(NewDeaths_ens, 0.95, 3);

fprintf('Simulation complete (%d-run ensemble).\n', Nrep);

%% 
%9 Derived quantities

I_total = Is_arr + Ia_arr;   % final representative replicate
I_nat   = sum(I_total, 2);

vacc_cov = (V1_arr + V2_arr + V3_arr) ./ repmat(Npop', Nsteps, 1);

for k = 1:NC
    c_obs = new_cases_sm(:,k);   c_mod = new_cases_model(:,k);
    d_obs = new_deaths_sm(:,k);  d_mod = new_deaths_model(:,k);
    rmse_c = sqrt(mean((c_mod - c_obs).^2));
    rmse_d = sqrt(mean((d_mod - d_obs).^2));
    if std(c_obs) > 0 && std(c_mod) > 0
        r_c = corr(c_mod, c_obs);
    else, r_c = NaN; end
    fprintf('  %4s: cases RMSE=%7.1f  r=%.3f  deaths RMSE=%.2f\n', ...
        STATES{k}, rmse_c, r_c, rmse_d);
end

%% 
%10 Calibration vs simulation diagnostic
fprintf('\nBuilding calibration-vs-simulation diagnostic...\n');
cal_cases_nat  = zeros(Nsteps, 1);
cal_deaths_nat = zeros(Nsteps, 1);
for k = 1:NC
    [~, moD] = calibration_residuals_1patch(theta_fit_d(k,:)', cal_d2_all{k});
    [~, moO] = calibration_residuals_1patch(theta_fit_o(k,:)', cal_o2_all{k});
    cal_cases_nat(win_d)  = cal_cases_nat(win_d)  + moD.cases;
    cal_cases_nat(win_o)  = cal_cases_nat(win_o)  + moO.cases;
    cal_deaths_nat(win_d) = cal_deaths_nat(win_d) + moD.deaths;
    cal_deaths_nat(win_o) = cal_deaths_nat(win_o) + moO.deaths;
end

sim_cases_nat  = sum(new_cases_model,  2);
sim_deaths_nat = sum(new_deaths_model, 2);
obs_cases_nat  = sum(new_cases_sm,  2);
obs_deaths_nat = sum(new_deaths_sm, 2);

om_switch = d0 + omicron_day - 1;
figure('Name','Calibration vs Simulation','Position',[60 60 1300 480]);
subplot(1,2,1); hold on;
plot(dates, obs_cases_nat, 'k--', 'LineWidth',1.4);
plot(dates, cal_cases_nat, 'b-',  'LineWidth',1.8);
plot(dates, sim_cases_nat, 'r-',  'LineWidth',1.6);
xline(om_switch, ':', 'Omicron');
legend('Observed (state sum)','Calibration fit (single-patch, no mobility)', ...
       'Simulation median (multi-patch, mobility)','Location','best');
title('National Daily Cases'); ylabel('cases/day'); grid on;
xlim(plot_xl);
subplot(1,2,2); hold on;
plot(dates, obs_deaths_nat, 'k--', 'LineWidth',1.4);
plot(dates, cal_deaths_nat, 'b-',  'LineWidth',1.8);
plot(dates, sim_deaths_nat, 'r-',  'LineWidth',1.6);
xline(om_switch, ':', 'Omicron');
legend('Observed (state sum)','Calibration fit','Simulation median','Location','best');
title('National Daily Deaths'); ylabel('deaths/day'); grid on;
xlim(plot_xl);
sgtitle('Calibration vs Simulation — where the two models diverge');

% Quantify the divergence over each window
gap_d = sum(sim_cases_nat(win_d)) / max(sum(cal_cases_nat(win_d)), 1);
gap_o = sum(sim_cases_nat(win_o)) / max(sum(cal_cases_nat(win_o)), 1);
fprintf('  Sim/Calib cumulative-case ratio:  Delta=%.2fx  Omicron=%.2fx\n', gap_d, gap_o);
fprintf('  (>1 => the simulation produces more cases than the calibration was fit to)\n');
fprintf('  Sim vs Observed cumulative-case ratio (Omicron): %.2fx\n', ...
        sum(sim_cases_nat(win_o)) / max(sum(obs_cases_nat(win_o)), 1));

% Per-state localisation
fprintf('  Per-state cumulative cases  [sim / calib / observed], Delta | Omicron:\n');
cal_cases_st = zeros(Nsteps, NC);
for k = 1:NC
    [~, moD] = calibration_residuals_1patch(theta_fit_d(k,:)', cal_d2_all{k});
    [~, moO] = calibration_residuals_1patch(theta_fit_o(k,:)', cal_o2_all{k});
    cal_cases_st(win_d,k) = moD.cases;
    cal_cases_st(win_o,k) = moO.cases;
    sD = sum(new_cases_model(win_d,k));  cD = sum(cal_cases_st(win_d,k));  oD = sum(new_cases_sm(win_d,k));
    sO = sum(new_cases_model(win_o,k));  cO = sum(cal_cases_st(win_o,k));  oO = sum(new_cases_sm(win_o,k));
    fprintf('    %-3s  D: sim=%8.0f calib=%8.0f obs=%8.0f (%.2fx) | O: sim=%9.0f calib=%9.0f obs=%9.0f (%.2fx)\n', ...
        STATES{k}, sD, cD, oD, sD/max(cD,1), sO, cO, oO, sO/max(cO,1));
end

%% 
%11 Figures Output
cmap = lines(NC);

figure('Name','Case Fit','Position',[40 40 1400 700]);
tl = tiledlayout(2,4,'Padding','compact','TileSpacing','compact');
title(tl,'Daily Cases: Data (red) vs Model median (blue), 5-95% band  |  7-day smoothed');
xlabel(tl,'Date');
for k = 1:NC
    nexttile;
    fill([dates, fliplr(dates)], [cases_q05(:,k)', fliplr(cases_q95(:,k)')], ...
         [0.80 0.80 1], 'EdgeColor','none', 'FaceAlpha',0.5); hold on;
    plot(dates, new_cases_sm(:,k),    'r--', 'LineWidth',0.8);
    plot(dates, new_cases_model(:,k), 'b',   'LineWidth',1.0);
    title(STATES{k}); ylabel('Cases Per Day');
    grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');
    if k == 1, legend('5-95% band','Data','Model','Location','best'); end
end

figure('Name','Death Fit','Position',[40 40 1400 700]);
tl2 = tiledlayout(2,4,'Padding','compact','TileSpacing','compact');
title(tl2,'Daily Deaths: Data (red) vs Model median (blue), 5-95% band  |  7-day smoothed');
xlabel(tl2,'Date');
for k = 1:NC
    nexttile;
    fill([dates, fliplr(dates)], [deaths_q05(:,k)', fliplr(deaths_q95(:,k)')], ...
         [0.80 0.80 1], 'EdgeColor','none', 'FaceAlpha',0.5); hold on;
    plot(dates, new_deaths_sm(:,k),    'r--', 'LineWidth',0.8);
    plot(dates, new_deaths_model(:,k), 'b',   'LineWidth',1.0);
    title(STATES{k}); ylabel('Deaths Per Day');
    grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');
    if k == 1, legend('5-95% band','Data','Model','Location','best'); end
end

figure('Name','>=1 Dose Coverage','Position',[40 40 1400 700]);
tl_v1 = tiledlayout(2,4,'Padding','compact','TileSpacing','compact');
title(tl_v1,'>=1 Dose Coverage (% 16+): Observed (blue) vs Low Hesitancy (red)');
xlabel(tl_v1,'Date');
for k = 1:NC
    nexttile;
    plot(dates, vacc_obs_v1(:,k)*100, 'b',   'LineWidth',1.3); hold on;
    plot(dates, min(cumV1_arr(:,k)/Npop16(k)*100, 100), 'r--','LineWidth',1.6);
    title(STATES{k}); ylabel('%');
    ylim([0 100]); grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');
    if k == 1, legend('Observed','Low Hesitancy','Location','best'); end
end

figure('Name','>=2 Dose Coverage','Position',[40 40 1400 700]);
tl_v2 = tiledlayout(2,4,'Padding','compact','TileSpacing','compact');
title(tl_v2,'>=2 Doses Coverage (% 16+): Observed (blue) vs Low Hesitancy (red)');
xlabel(tl_v2,'Date');
for k = 1:NC
    nexttile;
    plot(dates, vacc_obs_v2(:,k)*100, 'b',   'LineWidth',1.3); hold on;
    plot(dates, min(cumV2_arr(:,k)/Npop16(k)*100, 100), 'r--','LineWidth',1.6);
    title(STATES{k}); ylabel('%');
    ylim([0 100]); grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');
    if k == 1, legend('Observed','Low Hesitancy','Location','best'); end
end

figure('Name','>=3 Dose Coverage','Position',[40 40 1400 700]);
tl_v3 = tiledlayout(2,4,'Padding','compact','TileSpacing','compact');
title(tl_v3,'>=3 Doses (Booster) Coverage (% 16+): Observed (blue) vs Low Hesitancy (red)');
xlabel(tl_v3,'Date');
for k = 1:NC
    nexttile;
    plot(dates, vacc_obs_v3(:,k)*100, 'b',   'LineWidth',1.3); hold on;
    plot(dates, min(cumV3_arr(:,k)/Npop16(k)*100, 100), 'r--','LineWidth',1.6);
    title(STATES{k}); ylabel('%');
    ylim([0 100]); grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');
    if k == 1, legend('Observed','Low Hesitancy','Location','best'); end
end

figure('Name','National Infectious','Position',[40 40 1100 420]);
area(dates, I_total, 'LineStyle','none');
colororder(cmap);
legend(STATES,'Location','best','NumColumns',2);
xlabel('Date'); ylabel('Infectious persons');
title('Total Infectious by State (I_s + I_a)');
grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

phi_mean = mean(PHI, 3) * 1e4;
figure('Name','Mobility Matrix','Position',[40 40 620 520]);
imagesc(phi_mean); cb = colorbar; colormap(parula);
ylabel(cb, 'Travel rate (\times10^{-4} day^{-1})');
set(gca,'XTick',1:NC,'XTickLabel',STATES,'YTick',1:NC,'YTickLabel',STATES);
xlabel('Origin State'); ylabel('Destination State');
title('Mean Mobility Matrix (\times10^{-4} day^{-1})');
axis square; box on;

figure('Name','Beta(t)','Position',[40 40 900 360]);
colororder(cmap);
plot(dates, beta_t, 'LineWidth', 1.0);
legend(STATES,'Location','best','NumColumns',2);
xlabel('Date'); ylabel('\beta(t)');
title('Per-State Stochastic Transmission Rate \beta(t)  —  OU Process');
grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

figure('Name','Vaccine Hesitancy','Position',[40 40 1000 420]);
plot(dates, x_arr, 'LineWidth',1.4);
legend(STATES,'Location','best','NumColumns',2);
xlabel('Date'); ylabel('Vaccine Hesitancy (%)');
title('Vaccine Hesitancy by State — Replicator SDE driven by case/death salience');
ylim([0 1]); grid on; xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

%% 
%12 Scenario Analysis
fprintf('\nRunning scenario analysis...\n');

scenarios  = struct('name',{'Low Hesitancy (x=0.15)','Medium Hesitancy (x=0.30)', ...
                             'High Hesitancy (x=0.60)'}, ...
                    'xrefval',{0.15, 0.30, 0.60});   % Baseline
colors_sc  = {[0.2 0.5 0.8], [0.1 0.7 0.3], [0.8 0.2 0.2]};
Nrep_sc    = 50;   % replicates per scenario

figure('Name','Scenarios','Position',[40 40 1500 460]);
tl3 = tiledlayout(1,3,'Padding','compact','TileSpacing','compact');
title(tl3,'Scenario Analysis — National Totals by Hesitancy (ensemble median \pm 5-95%)');

hcase = gobjects(1,3);  hdeath = gobjects(1,3);  hcov = gobjects(1,3);
scen_cases_med  = zeros(Nsteps,3);   % per-scenario ensemble-median trajectories
scen_deaths_med = zeros(Nsteps,3);
scen_cov_med    = zeros(Nsteps,3);

% Per-scenario total cases/deaths over the whole window 
scen_total_cases_med  = zeros(3,1);  scen_total_cases_lo  = zeros(3,1);  scen_total_cases_hi  = zeros(3,1);
scen_total_deaths_med = zeros(3,1);  scen_total_deaths_lo = zeros(3,1);  scen_total_deaths_hi = zeros(3,1);

for sc_idx = 1:3
    hes_sc = hes;  hes_sc.x_ref = scenarios(sc_idx).xrefval;   % persistent disposition
    x0_sc  = scenarios(sc_idx).xrefval * ones(NC,1);           % start at the disposition

    % Vaccination 'warm-up' before time bounds
    warm_days = days(datetime('2021-06-01') - VACC_RAMP_START);   
    x_ref_sc  = scenarios(sc_idx).xrefval;
    g_wu      = vacc_uptake(x_ref_sc, vacc_rate_sens, vacc_x_hist);
    dc1 = vacc_obs_v1(1,:)'/warm_days;   
    dc2 = vacc_obs_v2(1,:)'/warm_days;
    dc3 = vacc_obs_v3(1,:)'/warm_days;
    c1 = zeros(NC,1); c2 = zeros(NC,1); c3 = zeros(NC,1);   
    for wd = 1:warm_days
        for k = 1:NC
            c1(k) = c1(k) + dc1(k)*g_wu*vacc_headroom(c1(k), x_ref_sc, vacc_r, vacc_w, vacc_C_hist(k,1), vacc_x_hist);
            c2(k) = c2(k) + dc2(k)*g_wu*vacc_headroom(c2(k), x_ref_sc, vacc_r, vacc_w, vacc_C_hist(k,2), vacc_x_hist);
            c3(k) = c3(k) + dc3(k)*g_wu*vacc_headroom(c3(k), x_ref_sc, vacc_r, vacc_w, vacc_C_hist(k,3), vacc_x_hist);
        end
    end
    c2 = min(c2, c1);  c3 = min(c3, c2);   % enforce the dose ladder
    V10_sc = round(max(c1 - c2, 0) .* Npop16);  
    V20_sc = round(max(c2 - c3, 0) .* Npop16);  
    V30_sc = round(c3 .* Npop16);               
    S0_sc  = max(Npop - E0 - Is0 - Ia0 - Q0 - R0v - V10_sc - V20_sc - V30_sc, 0);
    fprintf('  Warm-up %-26s: nat >=1 dose at June-1 = %.3f (observed %.3f)\n', ...
            scenarios(sc_idx).name, sum(V10_sc+V20_sc+V30_sc)/sum(Npop16), ...
            sum(V10+V20+V30)/sum(Npop16));

    Cases_rep  = zeros(Nsteps, Nrep_sc);
    Deaths_rep = zeros(Nsteps, Nrep_sc);
    Cov_rep    = zeros(Nsteps, Nrep_sc);

    for rep = 1:Nrep_sc
        rng(1000 + rep);   % common random numbers across scenarios; distinct per replicate

        S_sc=S0_sc; E_sc=E0; Is_sc=Is0; Ia_sc=Ia0; Q_sc=Q0;
        R_sc=R0v; V1_sc=V10_sc; V2_sc=V20_sc; V3_sc=V30_sc; x_sc=x0_sc;
        Sw_sc=Sw0;
        % Shadow chain
        Sc_sc = min(Nkids, S0_sc);   
        Rf_sc = zeros(NC,1);   
        xi_sc = zeros(NC,1);
        Mc_sc = 1e5.*max(p_d.*alpha_mat(1,:)'.*E_sc.*detect_rate_mat(1,:)', 0)./Npop;
        Md_sc = 1e5.*max(mu_d_d_vec.*(Is_sc+Q_sc), 0)./Npop;

        nat_cases_sc  = zeros(Nsteps, 1);
        nat_deaths_sc = zeros(Nsteps, 1);
        nat_cov_sc    = zeros(Nsteps, 1);
        nat_cov_sc(1) = sum(V1_sc + V2_sc + V3_sc);   % national >=1 dose at t=1

    for t = 1:Nsteps-1
        xi_sc = xi_sc - theta_ou*xi_sc*dt + sigma_b*sqrt(dt)*randn(NC,1);
        Phi   = PHI(:,:, min(t,Ndays));

        nS=S_sc; nE=E_sc; nIs=Is_sc; nIa=Ia_sc;
        nQ=Q_sc; nR=R_sc; nV1=V1_sc; nV2=V2_sc; nV3=V3_sc; nx=x_sc;
        nSw=Sw_sc;
        nEn=En_sc; nIsn=Isn_sc; nIan=Ian_sc; nQn=Qn_sc; nRn=Rn_sc;
        nSc=Sc_sc; nRf=Rf_sc;
        nMc=Mc_sc; nMd=Md_sc;
        dose1_inc = 0;

        for k = 1:NC
            S=S_sc(k); E=E_sc(k); Is=Is_sc(k); Ia=Ia_sc(k);
            Q=Q_sc(k); Rv=R_sc(k); V1=V1_sc(k); V2=V2_sc(k); V3=V3_sc(k);
            xk=x_sc(k); Sw=Sw_sc(k);
            En=En_sc(k); Isn=Isn_sc(k); Ian=Ian_sc(k); Qn=Qn_sc(k); Rn=Rn_sc(k);
            Sc=Sc_sc(k);
            Rf=Rf_sc(k);
            N = max(S+Sw+E+Is+Ia+Q+Rv+V1+V2+V3, 1);

            beta_now = beta_base_mat(t+1,k) * npi_scale(t+1,k) * exp(xi_sc(k)) ...
                       * behav_suppression(Mc_sc(k), Md_sc(k), hes_sc);
            mob_in   = sum(Phi(k,:).*(Is_sc'+Ia_sc')) - Phi(k,k)*(Is+Ia);
            lam      = beta_now*(Is+Ia)/N;
            lam_tot  = lam + beta_now*mob_in/N;

            alpha_k   = alpha_mat(t+1,k); gamma_s_k = gamma_s_mat(t+1,k);
            gamma_a_k = gamma_a_mat(t+1,k); gamma_q_k = gamma_q_mat(t+1,k);
            delta_k   = delta_mat(t+1,k); eps1_k    = eps1_mat(t+1,k);
            eps2_k    = eps2_mat(t+1,k);  eps3_k    = eps3_mat(t+1,k);
            omega_k   = omega_mat(t+1,k); p_k       = p_vec(t+1);
            mu_d_k    = mu_d_mat(t+1,k);

            l1=(1-eps1_k)*lam_tot; l2=(1-eps2_k)*lam_tot; l3=(1-eps3_k)*lam_tot;

            % Reinfection flow 
            Ru=max(0,Rv-Rf);
            lamR_fresh=lam_tot*(1-eps_R_fresh_vec(t+1));
            lamR_floor=lam_tot*(1-eps_R_floor_vec(t+1));
            reinf_tot=lamR_fresh*Rf+lamR_floor*Ru;
            reinf_naive=reinf_tot*(Rn/max(Rv,1));

            mu_d_eff = death_rate_eff(mu_d_k, lam_tot*(S+Sw), l1*V1, l2*V2, l3*V3, ...
                                      [epsD1_vec(t+1); epsD2_vec(t+1); epsD3_vec(t+1)], ...
                                      reinf_tot, eps_R_death);

            VC=(V1+V2+V3)/N;
            cases_now  = p_k*alpha_k*E*detect_rate_mat(t,k);
            deaths_now = mu_d_eff*(Is+Q);
            [nx(k), nMc(k), nMd(k)] = hesitancy_step(xk, cases_now, deaths_now, ...
                VC, N, Mc_sc(k), Md_sc(k), sqrt(dt)*randn, hes_sc);

            x_disp=hes_sc.x_ref;   % scenario disposition drives the ceiling/rate 
            g_up=vacc_uptake(x_disp,vacc_rate_sens,vacc_x_hist);   % demand speed multiplier
            nu1e=nu_dyn(1,t,k)*g_up*vacc_headroom((V1+V2+V3)/Npop16(k),x_disp,vacc_r,vacc_w,vacc_C_hist(k,1),vacc_x_hist);
            nu2e=nu_dyn(2,t,k)*g_up*vacc_headroom((V2+V3)/Npop16(k),x_disp,vacc_r,vacc_w,vacc_C_hist(k,2),vacc_x_hist);
            nu3e=nu_dyn(3,t,k)*g_up*vacc_headroom(V3/Npop16(k),x_disp,vacc_r,vacc_w,vacc_C_hist(k,3),vacc_x_hist);
            Se = max(0, S - Sc);   % eligible (16+) susceptibles: first doses only
            dose1_inc = dose1_inc + nu1e*Se*dt;   % national new first doses this step

            Rw=max(0,Rv-Rn);   % recovered who had been vaccinated before infection
            dS=(Lambda(k)-lam_tot*S-nu1e*Se-mu*S+omega_R*Rn)*dt;
            dSc=(Lambda(k)-lam_tot*Sc-mu*Sc)*dt;   % under-16 subset of S 
            dSw=(omega_k*V3-lam_tot*Sw-mu*Sw+omega_R*Rw)*dt;   % booster + vacc-recovered waning
            dE=(lam_tot*S+lam_tot*Sw+reinf_tot+l1*V1+l2*V2+l3*V3-alpha_k*E-mu*E)*dt;
            dIs=(p_k*alpha_k*E-(gamma_s_k+delta_k+mu+mu_d_eff)*Is)*dt;
            dIa=((1-p_k)*alpha_k*E-(gamma_a_k+mu)*Ia)*dt;
            dQ=(delta_k*Is-(gamma_q_k+mu+mu_d_eff)*Q)*dt;
            dR=(gamma_s_k*Is+gamma_a_k*Ia+gamma_q_k*Q-mu*Rv-omega_R*Rv-reinf_tot)*dt;
            dRf=(gamma_s_k*Is+gamma_a_k*Ia+gamma_q_k*Q-mu*Rf-omega_R*Rf-omega_NI*Rf-lamR_fresh*Rf)*dt;

            % Shadow chain
            dEn=(lam_tot*S+reinf_naive-alpha_k*En-mu*En)*dt;
            dIsn=(p_k*alpha_k*En-(gamma_s_k+delta_k+mu+mu_d_eff)*Isn)*dt;
            dIan=((1-p_k)*alpha_k*En-(gamma_a_k+mu)*Ian)*dt;
            dQn=(delta_k*Isn-(gamma_q_k+mu+mu_d_eff)*Qn)*dt;
            dRn=(gamma_s_k*Isn+gamma_a_k*Ian+gamma_q_k*Qn-mu*Rn-omega_R*Rn-reinf_naive)*dt;
            dV1=(nu1e*Se-nu2e*V1-mu*V1)*dt;
            dV2=(nu2e*V1-nu3e*V2-mu*V2)*dt;
            dV3=(nu3e*V2-omega_k*V3-mu*V3)*dt;

            nS(k)=max(0,S+dS); nE(k)=max(0,E+dE);
            nSw(k)=max(0,Sw+dSw);
            nIs(k)=max(0,Is+dIs); nIa(k)=max(0,Ia+dIa);
            nQ(k)=max(0,Q+dQ); nR(k)=max(0,Rv+dR);
            nV1(k)=max(0,V1+dV1); nV2(k)=max(0,V2+dV2); nV3(k)=max(0,V3+dV3);
       
            nEn(k)=min(max(0,En+dEn),nE(k));   nIsn(k)=min(max(0,Isn+dIsn),nIs(k));
            nIan(k)=min(max(0,Ian+dIan),nIa(k)); nQn(k)=min(max(0,Qn+dQn),nQ(k));
            nRn(k)=min(max(0,Rn+dRn),nR(k));
            nSc(k)=min(max(0,Sc+dSc),nS(k)); 
            nRf(k)=min(max(0,Rf+dRf),nR(k));   

            nat_cases_sc(t)  = nat_cases_sc(t)  + p_k*alpha_k*E*detect_rate_mat(t,k);
            nat_deaths_sc(t) = nat_deaths_sc(t) + mu_d_eff*(Is+Q);
        end
        S_sc=nS; E_sc=nE; Is_sc=nIs; Ia_sc=nIa;
        Q_sc=nQ; R_sc=nR; V1_sc=nV1; V2_sc=nV2; V3_sc=nV3; x_sc=nx;
        Sw_sc=nSw;
        En_sc=nEn; Isn_sc=nIsn; Ian_sc=nIan; Qn_sc=nQn; Rn_sc=nRn;
        Sc_sc=nSc; Rf_sc=nRf;
        Mc_sc=nMc; Md_sc=nMd;
        nat_cov_sc(t+1) = nat_cov_sc(t) + dose1_inc;
        end   % for t

        % Boundary fix 
        nat_cases_sc(Nsteps)  = nat_cases_sc(Nsteps-1);
        nat_deaths_sc(Nsteps) = nat_deaths_sc(Nsteps-1);

        Cases_rep(:,rep)  = nat_cases_sc;
        Deaths_rep(:,rep) = nat_deaths_sc;
        Cov_rep(:,rep)    = min(nat_cov_sc/sum(Npop16)*100, 100);
    end   % for rep

    % Total cases/deaths over the whole window, per replicate, then summarised
    tot_cases_rep  = sum(Cases_rep, 1)';
    tot_deaths_rep = sum(Deaths_rep, 1)';
    scen_total_cases_med(sc_idx)  = median(tot_cases_rep);
    scen_total_cases_lo(sc_idx)   = quantile(tot_cases_rep, 0.05);
    scen_total_cases_hi(sc_idx)   = quantile(tot_cases_rep, 0.95);
    scen_total_deaths_med(sc_idx) = median(tot_deaths_rep);
    scen_total_deaths_lo(sc_idx)  = quantile(tot_deaths_rep, 0.05);
    scen_total_deaths_hi(sc_idx)  = quantile(tot_deaths_rep, 0.95);

    % Ensemble summaries
    cases_med  = movmean(median(Cases_rep,2),7);
    cases_lo  = movmean(quantile(Cases_rep,0.05,2),7);
    cases_hi  = movmean(quantile(Cases_rep,0.95,2),7);
    deaths_med = movmean(median(Deaths_rep,2),7);
    deaths_lo = movmean(quantile(Deaths_rep,0.05,2),7);
    deaths_hi = movmean(quantile(Deaths_rep,0.95,2),7);
    cov_med    = median(Cov_rep,2);
    cov_lo    = quantile(Cov_rep,0.05,2);
    cov_hi    = quantile(Cov_rep,0.95,2);

    scen_cases_med(:,sc_idx)  = cases_med;
    scen_deaths_med(:,sc_idx) = deaths_med;
    scen_cov_med(:,sc_idx)    = cov_med;

    cc = colors_sc{sc_idx};
    nexttile(1);
    fill([dates, fliplr(dates)], [cases_lo', fliplr(cases_hi')], cc, ...
         'FaceAlpha',0.15,'EdgeColor','none'); hold on;
    hcase(sc_idx) = plot(dates, cases_med, 'Color',cc,'LineWidth',1.8);
    nexttile(2);
    fill([dates, fliplr(dates)], [deaths_lo', fliplr(deaths_hi')], cc, ...
         'FaceAlpha',0.15,'EdgeColor','none'); hold on;
    hdeath(sc_idx) = plot(dates, deaths_med, 'Color',cc,'LineWidth',1.8);
    nexttile(3);
    fill([dates, fliplr(dates)], [cov_lo', fliplr(cov_hi')], cc, ...
         'FaceAlpha',0.15,'EdgeColor','none'); hold on;
    hcov(sc_idx) = plot(dates, cov_med, 'Color',cc,'LineWidth',1.8);
end

nexttile(1);
hobs1 = plot(dates, nat_cases_obs_sm, 'k--','LineWidth',1.0);
legend([hcase, hobs1], [{scenarios.name}, {'Observed'}],'Location','best'); grid on;
xlabel('Date'); ylabel('New Cases Per Day'); title('National Daily Cases');
xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

nexttile(2);
hobs2 = plot(dates, nat_deaths_obs_sm, 'k--','LineWidth',1.0);
legend([hdeath, hobs2], [{scenarios.name}, {'Observed'}],'Location','best'); grid on;
xlabel('Date'); ylabel('New Deaths Per Day'); title('National Daily Deaths');
xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

nat_vacc_obs_v1 = sum(vacc_obs_v1 .* repmat(Npop16', Nsteps, 1), 2) / sum(Npop16) * 100;
nexttile(3);
hobs3 = plot(dates, nat_vacc_obs_v1, 'k--','LineWidth',1.0);
legend([hcov, hobs3], [{scenarios.name}, {'Observed'}],'Location','best'); grid on;
xlabel('Date'); ylabel('1st Dose Coverage (% 16+)');
title('National Vaccination Coverage');
ylim([0 100]); xlim(plot_xl); datetick('x','mmm-yy','keeplimits');

% Additional results
commify = @(x) regexprep(sprintf('%.0f', x), '\d(?=(\d{3})+(?!\d))', '$0,');
obs_total_cases  = sum(nat_cases_obs);
obs_total_deaths = sum(nat_deaths_obs);

figure('Name','Scenario Totals','Position',[40 40 820 260]);
ax = axes('Position',[0.04 0.06 0.92 0.78]); axis(ax,'off');
xlim(ax,[0 1]); ylim(ax,[0 1]); hold(ax,'on');
title(ax, sprintf('Scenario Totals (%s \\rightarrow %s)', ...
      datestr(dates(1),'mmm yyyy'), datestr(dates(end),'mmm yyyy')), ...
      'FontWeight','bold','FontSize',12);

col_x  = [0.02, 0.38, 0.70];
n_rows = 3 + 1 + 1;   % 3 scenarios + observed + header
row_y  = linspace(0.90, 0.06, n_rows);

text(ax, col_x(1), row_y(1), 'Scenario',                        'FontWeight','bold','FontSize',10);
text(ax, col_x(2), row_y(1), 'Total cases (median [5-95%])',     'FontWeight','bold','FontSize',10);
text(ax, col_x(3), row_y(1), 'Total deaths (median [5-95%])',    'FontWeight','bold','FontSize',10);
line(ax, [0 1], mean(row_y(1:2))*[1 1], 'Color',[0.3 0.3 0.3], 'LineWidth',1);

for sc_idx = 1:3
    ry = row_y(sc_idx+1);
    text(ax, col_x(1), ry, scenarios(sc_idx).name, ...
         'Color', colors_sc{sc_idx}, 'FontWeight','bold','FontSize',10);
    text(ax, col_x(2), ry, sprintf('%s   [%s - %s]', commify(scen_total_cases_med(sc_idx)), ...
         commify(scen_total_cases_lo(sc_idx)), commify(scen_total_cases_hi(sc_idx))), 'FontSize',10);
    text(ax, col_x(3), ry, sprintf('%s   [%s - %s]', commify(scen_total_deaths_med(sc_idx)), ...
         commify(scen_total_deaths_lo(sc_idx)), commify(scen_total_deaths_hi(sc_idx))), 'FontSize',10);
end

line(ax, [0 1], mean(row_y(4:5))*[1 1], 'Color',[0.8 0.8 0.8], 'LineWidth',0.75);
text(ax, col_x(1), row_y(5), 'Observed (AU)', 'FontAngle','italic','FontSize',10);
text(ax, col_x(2), row_y(5), commify(obs_total_cases),  'FontAngle','italic','FontSize',10);
text(ax, col_x(3), row_y(5), commify(obs_total_deaths), 'FontAngle','italic','FontSize',10);

%% 
%13 Save results
results.dates             = dates;
results.states            = STATES;
results.Npop              = Npop;
results.new_cases_obs     = new_cases;
results.new_deaths_obs    = new_deaths;
results.new_cases_model   = new_cases_model;   % ensemble median
results.new_deaths_model  = new_deaths_model;  % ensemble median
results.cases_q05         = cases_q05;
results.cases_q95         = cases_q95;
results.deaths_q05        = deaths_q05;
results.deaths_q95        = deaths_q95;
results.Nrep              = Nrep;
results.S  = S_arr;  results.E  = E_arr;  results.Is = Is_arr;
results.Ia = Ia_arr; results.Q  = Q_arr;  results.R  = R_arr;
results.V1 = V1_arr; results.V2 = V2_arr; results.V3 = V3_arr;
results.Sw = Sw_arr;
results.Rn = Rn_arr;             
results.Rw = max(0, R_arr - Rn_arr);  
results.Sc = Sc_arr;                  
results.Se = max(0, S_arr - Sc_arr); 
results.x  = x_arr;  results.beta_t = beta_t;  results.phi0 = phi0;
results.Mc = Mc_arr; results.Md = Md_arr;  results.hes = hes;
results.mu_d_d_state = mu_d_d_vec;  
results.mu_d_o_state = mu_d_o_vec;  
results.scen_names       = {scenarios.name};
results.scen_xref        = [scenarios.xrefval];     
results.scen_cases_med    = scen_cases_med;            
results.scen_deaths_med   = scen_deaths_med;            
results.scen_cov_med      = scen_cov_med;             
results.scen_total_cases_med  = scen_total_cases_med;  results.scen_total_cases_lo  = scen_total_cases_lo;  results.scen_total_cases_hi  = scen_total_cases_hi;
results.scen_total_deaths_med = scen_total_deaths_med; results.scen_total_deaths_lo = scen_total_deaths_lo; results.scen_total_deaths_hi = scen_total_deaths_hi;
results.obs_total_cases  = obs_total_cases;
results.obs_total_deaths = obs_total_deaths;
results.theta_delta   = theta_fit_d;   
results.theta_omicron = theta_fit_o; 
results.theta_names   = {'beta0','alpha','gamma_s','gamma_a','gamma_q','delta', ...
                          'epsilon1','epsilon2','epsilon3','omega'};
results.fixed = struct('p_d',p_d,'p_o',p_o, ...
                        'mu_d_d',mu_d_d,'mu_d_o',mu_d_o, ...
                        'CFR_d',CFR_d,'CFR_o',CFR_o, ...
                        'beta0_d_guess',beta0_d_guess,'beta0_o_guess',beta0_o_guess, ...
                        'R0_d',R0_d,'R0_o',R0_o,'jensen_corr',jensen_corr, ...
                        'Reff_pre',Reff_pre, ...
                        'epsD_d',epsD_d,'epsD_o',epsD_o, ...
                        'vacc_r',vacc_r,'vacc_w',vacc_w,'w_death',w_death, ...
                        'vacc_rate_sens',vacc_rate_sens,'vacc_x_hist',vacc_x_hist);
results.fixed.beta0_pre  = beta0_pre;
results.fixed.NPI_SOURCE = NPI_SOURCE;
results.fixed.NPI_POWER  = NPI_POWER;
results.npi_scale = npi_scale;   
results.npi_step  = npi_step;    

save('MultiCity_Results_PerState.mat', 'results');
fprintf('\nResults saved to MultiCity_Results_PerState.mat\n');

%% 
%14 Sensitivity
sens_om_ratio = sum(sim_cases_nat(win_o)) / max(sum(obs_cases_nat(win_o)), 1);
[sens_pk, sens_pk_i] = max(sum(new_cases_model, 2));
sens_tail     = mean(sum(new_cases_model(end-29:end, :), 2));   % mean daily, last 30d
sens_row = table( ...
    string(SENS_VARIANT), eps_R_floor_o, 1/omega_NI, 1/omega_R, eps_R_death, hes.b_max, ...
    hes.ab_c/hes.ab_d, hes.tau_c, theta_ou, ...
    sens_om_ratio, sens_pk, string(datestr(dates(sens_pk_i),'dd-mmm-yy')), sens_tail, ...
    scen_total_cases_med(1),  scen_total_deaths_med(1), ...
    scen_total_cases_med(3)/max(scen_total_cases_med(1),1), ...
    scen_total_deaths_med(3)/max(scen_total_deaths_med(1),1), ...
    'VariableNames', {'variant','epsRfloor_o','omegaNI_days','omegaR_days','epsRdeath','b_max', ...
    'ab_c_over_ab_d','tau_days','theta_ou','omicron_sim_over_obs','natl_peak', ...
    'peak_date','tail30_mean','base_cases','base_deaths', ...
    'highhes_case_ratio','highhes_death_ratio'});

% Append this run's row. Guard against a STALE accumulator
if exist('SENS_RESULTS','var') && ~isempty(SENS_RESULTS) && ...
        isequal(SENS_RESULTS.Properties.VariableNames, sens_row.Properties.VariableNames)
    SENS_RESULTS = [SENS_RESULTS; sens_row];   
else
    if exist('SENS_RESULTS','var') && ~isempty(SENS_RESULTS)
        warning('SENS_RESULTS schema changed since the last run; starting a fresh accumulation.');
    end
    SENS_RESULTS = sens_row;
end
writetable(SENS_RESULTS, 'Sensitivity_Summary.csv');
fprintf('\n===== SENSITIVITY SUMMARY (accumulated across runs) =====\n');
disp(SENS_RESULTS);
fprintf('Saved to Sensitivity_Summary.csv\n');
fprintf(['Read the last two columns first: if highhes_case_ratio / highhes_death_ratio\n' ...
         'are stable across variants, the hesitancy conclusion is robust to the\n' ...
         'parameters that evidence does not pin.\n']);

%%
% Local functions

function [res, mod_out] = calibration_residuals_1patch(theta, d)

beta0    = theta(1);
alpha    = theta(2);  gamma_s  = theta(3);  gamma_a = theta(4);
gamma_q  = theta(5);  delta    = theta(6);
epsilon1 = theta(7);  epsilon2 = theta(8);  epsilon3 = theta(9);
omega    = theta(10);
detect_k = theta(11);   

p=d.p; mu_d=d.mu_d; mu=d.mu; omega_R=d.omega_R;
Lambda=d.Lambda; x_min=d.x_min; Ndays=d.Ndays; dt=1;

Sk=d.S0; Ek=d.E0; Isk=d.Is0; Iak=d.Ia0;
Qk=d.Q0; Rvk=d.R0v; V1k=d.V10; V2k=d.V20; V3k=d.V30; xk=d.x0;
Swk=d.Sw0;   
% Shadow chain
Enk=d.En0; Isnk=d.Isn0; Iank=d.Ian0; Qnk=d.Qn0; Rnk=d.Rn0;
Sck = min(d.N - d.N16, Sk);  
Rfk = min(d.Rf0, Rvk);  

mod_cases    = zeros(Ndays, 1);
mod_cases(1) = p * alpha * Ek;
mod_deaths    = zeros(Ndays, 1);
mod_deaths(1) = mu_d * (Isk+Qk);

% Risk-salience EMA memory
Mc = 1e5*max(p*alpha*Ek*detect_k,0)/d.N;
Md = 1e5*max(mu_d*(Isk+Qk),0)/d.N;

for t = 1:Ndays-1
    Nk    = max(Sk+Swk+Ek+Isk+Iak+Qk+Rvk+V1k+V2k+V3k, 1);
    lam   = beta0 * d.npi_scale(t) * behav_suppression(Mc, Md, d.hes) * (Isk+Iak) / Nk;
    l1    = (1 - epsilon1) * lam;
    l2    = (1 - epsilon2) * lam;
    l3    = (1 - epsilon3) * lam;

    % Reinfection flow
    Ruk = max(0, Rvk - Rfk);
    lamR_fresh = lam * (1 - d.eps_R_fresh);
    lamR_floor = lam * (1 - d.eps_R_floor);
    reinf_tot   = lamR_fresh*Rfk + lamR_floor*Ruk;
    reinf_naive = reinf_tot * (Rnk / max(Rvk,1));

    mu_d_eff = death_rate_eff(mu_d, lam*(Sk+Swk), l1*V1k, l2*V2k, l3*V3k, d.epsD, ...
                              reinf_tot, d.eps_R_death);

    x_disp = d.hes.x_ref;   
    g_up = vacc_uptake(x_disp, d.vacc_rate_sens, d.vacc_x_hist);  
    nu1e = d.nu_dyn(1,t) * g_up * vacc_headroom((V1k+V2k+V3k)/d.N16, x_disp, d.vacc_r, d.vacc_w, d.C_hist(1), d.vacc_x_hist);
    nu2e = d.nu_dyn(2,t) * g_up * vacc_headroom((V2k+V3k)/d.N16,     x_disp, d.vacc_r, d.vacc_w, d.C_hist(2), d.vacc_x_hist);
    nu3e = d.nu_dyn(3,t) * g_up * vacc_headroom(V3k/d.N16,           x_disp, d.vacc_r, d.vacc_w, d.C_hist(3), d.vacc_x_hist);

    vacc_cov   = (V1k+V2k+V3k) / Nk;
    cases_now  = p*alpha*Ek*detect_k;
    deaths_now = mu_d_eff*(Isk+Qk);
    [xk, Mc, Md] = hesitancy_step(xk, cases_now, deaths_now, vacc_cov, ...
                                  d.N, Mc, Md, 0, d.hes);
    mod_deaths(t+1) = deaths_now;

    Rwk = max(0, Rvk - Rnk);   
    Sek = max(0, Sk - Sck);  
    dSc = (Lambda - lam*Sck - mu*Sck) * dt; 
    dS  = (Lambda - lam*Sk - nu1e*Sek - mu*Sk + omega_R*Rnk) * dt;
    dSw = (omega*V3k - lam*Swk - mu*Swk + omega_R*Rwk) * dt;  
    dE  = (lam*Sk + lam*Swk + reinf_tot + l1*V1k + l2*V2k + l3*V3k - alpha*Ek - mu*Ek) * dt;
    dIs = (p*alpha*Ek - (gamma_s+delta+mu+mu_d_eff)*Isk) * dt;
    dIa = ((1-p)*alpha*Ek - (gamma_a+mu)*Iak) * dt;
    dQ  = (delta*Isk - (gamma_q+mu+mu_d_eff)*Qk) * dt;
    dR  = (gamma_s*Isk + gamma_a*Iak + gamma_q*Qk - mu*Rvk - omega_R*Rvk - reinf_tot) * dt;
    dRf = (gamma_s*Isk + gamma_a*Iak + gamma_q*Qk - mu*Rfk - omega_R*Rfk - d.omega_NI*Rfk - lamR_fresh*Rfk) * dt;

    % Shadow chain
    dEn  = (lam*Sk + reinf_naive - alpha*Enk - mu*Enk) * dt;
    dIsn = (p*alpha*Enk - (gamma_s+delta+mu+mu_d_eff)*Isnk) * dt;
    dIan = ((1-p)*alpha*Enk - (gamma_a+mu)*Iank) * dt;
    dQn  = (delta*Isnk - (gamma_q+mu+mu_d_eff)*Qnk) * dt;
    dRn  = (gamma_s*Isnk + gamma_a*Iank + gamma_q*Qnk - mu*Rnk - omega_R*Rnk - reinf_naive) * dt;
    dV1 = (nu1e*Sek - nu2e*V1k - mu*V1k) * dt;
    dV2 = (nu2e*V1k - nu3e*V2k - mu*V2k) * dt;
    dV3 = (nu3e*V2k - omega*V3k - mu*V3k) * dt;

    Sk  = max(0, Sk+dS);   Ek  = max(0, Ek+dE);
    Swk = max(0, Swk+dSw);
    Isk = max(0, Isk+dIs); Iak = max(0, Iak+dIa);
    Qk  = max(0, Qk+dQ);   Rvk = max(0, Rvk+dR);
    V1k = max(0, V1k+dV1); V2k = max(0, V2k+dV2); V3k = max(0, V3k+dV3);
    Enk  = min(max(0, Enk+dEn),   Ek);   Isnk = min(max(0, Isnk+dIsn), Isk);
    Iank = min(max(0, Iank+dIan), Iak);  Qnk  = min(max(0, Qnk+dQn),   Qk);
    Rnk  = min(max(0, Rnk+dRn),   Rvk);
    Sck  = min(max(0, Sck+dSc),   Sk);   
    Rfk  = min(max(0, Rfk+dRf),   Rvk);

    mod_cases(t+1) = p * alpha * Ek;
end

skip  = 14;
obs_c = d.new_cases_sm(skip+1:end);
mod_c = mod_cases(skip+1:end) .* detect_k;   
c0    = max(0.05*mean(obs_c), 5);
res_c = log(max(mod_c,0) + c0) - log(obs_c + c0);

% Deaths residual 
obs_d = d.new_deaths_sm(skip+1:end);
mod_d = mod_deaths(skip+1:end);
c0_d  = max(0.05*mean(obs_d), 0.5);
res_d = log(max(mod_d,0) + c0_d) - log(obs_d + c0_d);

% Cumulative-incidence
res_cum = d.w_cum * (log(sum(max(mod_c,0)) + c0) - log(sum(obs_c) + c0));

res = [res_c; d.w_death * res_d; res_cum];

if nargout > 1
    mod_out = struct('cases', mod_cases .* detect_k, 'deaths', mod_deaths);
end
end


function [Sk,Ek,Isk,Iak,V1k,V2k,V3k] = extract_ics(day_idx, k, p_frac, ...
    new_cases_sm, detect_rate_val, vacc_obs_v1, vacc_obs_v2, vacc_obs_v3, Npop16, Npop)
% Extract initial conditions for state k at a given day from observed data
    I0k  = max(new_cases_sm(day_idx, k) / detect_rate_val, 0);
    Ek   = 2.0 * I0k;
    Isk  = p_frac * I0k;
    Iak  = (1-p_frac) * I0k;
    V1k  = max((vacc_obs_v1(day_idx,k) - vacc_obs_v2(day_idx,k)) * Npop16(k), 0);
    V2k  = max((vacc_obs_v2(day_idx,k) - vacc_obs_v3(day_idx,k)) * Npop16(k), 0);
    V3k  = max(vacc_obs_v3(day_idx,k) * Npop16(k), 0);
    Sk   = max(Npop(k) - Ek - Isk - Iak - V1k - V2k - V3k, 0);
end

function tbl = load_covid_xlsx(fpath, col_names)
    opts = detectImportOptions(fpath);
    opts.VariableNamesRange = '2:2';
    opts.DataRange          = 'A3';
    opts = setvartype(opts, opts.VariableNames{1}, 'char');
    opts = setvartype(opts, opts.VariableNames{2}, 'char');
    tbl  = readtable(fpath, opts);

    if nargin >= 2 && ~isempty(col_names)
        n = min(numel(col_names), width(tbl));
        tbl.Properties.VariableNames(1:n) = col_names(1:n);
    end

    tbl.date         = datetime(string(tbl.date), 'InputFormat','yyyy-MM-dd');
    tbl.location_key = string(tbl.location_key);

    for vi = 3:width(tbl)
        col = tbl{:, vi};
        if iscell(col)
            col = str2double(col);
            tbl{:, vi} = col;
        end
        if isnumeric(col)
            col(isnan(col)) = 0;
            tbl{:, vi} = col;
        end
    end
end


function [x_new, Mc_new, Md_new] = hesitancy_step(xk, cases_now, deaths_now, ...
        vacc_cov, N, Mc, Md, dW, hes)
% Replicator-diffusion hesitancy update 
    dt = hes.dt;
    y_c = 1e5 * max(cases_now, 0)  / max(N, 1);
    y_d = 1e5 * max(deaths_now, 0) / max(N, 1);
    A   = risk_alarm(Mc, Md, hes);
    c_v = max(hes.c_v_min, hes.c_v0 * (1 - hes.g_c*A));
    wI  = hes.wI0 * (1 + hes.g_w*A);                   
    % Replicator selection 
    gx    = xk * (1 - xk);
    sel   = hes.kappa * gx * (c_v - wI*A + hes.wC*vacc_cov - hes.b0);
    relax = hes.rho * (hes.x_ref - xk);
    x_new = max(hes.x_min, min(1, xk + (sel + relax)*dt + hes.sigma_x*gx*dW));
    Mc_new = Mc + (y_c - Mc)/hes.tau_c * dt;
    Md_new = Md + (y_d - Md)/hes.tau_d * dt;
end

function A = risk_alarm(Mc, Md, hes, a_c, a_d)
    if nargin < 5, a_c = hes.a_c; a_d = hes.a_d; end
    s_c = Mc / (Mc + hes.Kc);
    s_d = Md / (Md + hes.Kd);
    A   = (a_c*s_c + a_d*s_d) / (a_c + a_d);
end

function phi = behav_suppression(Mc, Md, hes)
    phi = 1 - hes.b_max * risk_alarm(Mc, Md, hes, hes.ab_c, hes.ab_d);
end

function mde = death_rate_eff(mu_d, infS, infV1, infV2, infV3, epsD, infR, epsR_death)
% Infection-weighted effective death rate
    inf_tot = infS + infV1 + infV2 + infV3 + infR;
    prot    = (epsD(1)*infV1 + epsD(2)*infV2 + epsD(3)*infV3 + epsR_death*infR) / max(inf_tot, 1e-12);
    mde     = mu_d * (1 - prot);
end


function h = vacc_headroom(cov, x, r, w, C_hist, x_hist)
% Willing-coverage-ceiling throttle on the vaccination rate
    C = min(1, max(0, C_hist - r*(x - x_hist)));
    h = max(0, min(1, (C - cov) / w));
end


function g = vacc_uptake(x, sens, x_hist)
% Hesitancy-dependent demand multiplier
    g = max(0, (1 - sens*x) / (1 - sens*x_hist));
end
