# Mangrove carbon stock: plot data to Mg C per hectare

**Illustrative data.** `carbon_stock.py` generates its own inventory from a fixed seed (8 circular plots of
7 m radius, 578 stems, one soil core per plot in 4 layers to 1 m). Nothing here is a measurement from a real
site.

## Method

1. **Stems to biomass**, Komiyama et al. (2005) common equations, D in cm, wood density in g/cm3:
   above-ground `0.251 * rho * D^2.46`, roots `0.199 * rho^0.899 * D^2.22` (kg per stem).
2. **Biomass to carbon** with a carbon fraction of 0.47 (IPCC 2006 default), CO2e = C x 44/12.
3. **Soil**: per layer, bulk density (g/cm3) x thickness (cm) x organic carbon (%) gives Mg C per ha;
   layers summed to 100 cm. Organic carbon from loss-on-ignition with two published equations:
   conventional `%Corg = 0.415 x %LOI + 2.89` (Palau mangroves) and localized `%OC = 0.36 x %LOI + 2.40`
   (Philippine mangroves), both as given in Salmo et al. (2024).
4. **Per hectare**: plot means, 90 % and 95 % confidence intervals (t distribution), coefficient of variation.
5. **Plots needed**: the smallest n with `t(n-1) x CV / sqrt(n)` within the target precision at 90 % confidence
   (simple random sampling).
6. **Survival lever**: kg CO2e per seedling planted for three diameters across survival rates.

## Results on the illustrative data

| | Value |
|---|---|
| Tree carbon (above-ground + roots) | 72.3 Mg C/ha, CV 35 %, 90 % CI +/- 23 % of the mean |
| Soil carbon to 1 m, conventional equation | 530.5 Mg C/ha (88 % of the total) |
| Soil carbon to 1 m, localized equation | 452.0 Mg C/ha |
| Gap caused by the equation choice alone | 78.5 Mg C/ha (14.8 %), more than all tree carbon |
| Plots needed for tree carbon at 90 % confidence | +/- 10 %: 35, +/- 15 %: 17, +/- 20 %: 11 |
| CO2e per surviving tree at 4 / 6 / 8 cm diameter | 16.6 / 43.4 / 86.0 kg |

For change between two campaigns the intercepts cancel but the slopes still differ by 13 % (0.36 / 0.415),
so a project should fix one equation in its SOP and calibrate it against elemental analysis of a subset.

## Placeholders

The two wood densities (Rhizophora mucronata 0.80, Avicennia marina 0.65 g/cm3) are placeholders: replace them
with species values from the Global Wood Density Database (Zanne et al. 2009) or a site study. The script prints
how many placeholders are in use every time it runs.

## Sources

- Komiyama, A., Poungparn, S. and Kato, S. (2005). Common allometric equations for estimating the tree weight
  of mangroves. Journal of Tropical Ecology 21: 471-477.
- Salmo et al. (2024). Improving soil carbon estimates of Philippine mangroves using localized organic matter to
  organic carbon equations. Carbon Balance and Management. doi:10.1186/s13021-024-00276-y
- IPCC (2006). Guidelines for National Greenhouse Gas Inventories, Volume 4.

## Files

`carbon_stock.py` (calculation), `charts.py` (figures), `outputs/` (CSV inputs and results, `summary.json`,
four PNG charts).
