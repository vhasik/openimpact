# openIMPACT

openIMPACT is an initiative to create a comprehensive, open-access environmental impact database focusing on high-impact building materials. Initially centered on Global Warming Potential (GWP) for North American products, the foundational models are adaptable for various regions and impact categories.

## Overview

- **EC3's EPD Database**: While EC3's EPD database is ideal for understanding the impact of individually procured products, it's not the best fit for early design due to its non-market-weighted nature and the potential lack of sufficient EPDs to capture the full range of impacts in a specific product market.

- **Need for openIMPACT**: There's a pressing need for transparent, high-quality, and free impact data for high-impact building materials that encapsulates the variation in impacts across markets, irrespective of the availability of product-specific EPDs.

- **Probabilistic Models**: Unlike traditional LCA models that rely on averages, openIMPACT employs probabilistic models. The results capture supply chain scenario probabilities based on market data and dependencies of primary parameters in manufacturing and supply chains. This is achieved by enhancing existing LCA models and leveraging the openIMPACT Monte Carlo Algorithm, which generates impacts for one realistic production (A1-A3) scenario at a time.

- **Methodology**: The project focuses on more robust estimation of A1-A3 (i.e., production) impacts using a Monte Carlo approach, but it also aims to provide data and models for other life cycle stages.

## Applications

- **Building LCA Tools**: The generated environmental impact data, along with uncertainty ranges, can be integrated into building LCA tools as generic data.

- **Parameterized Models**: These models can be employed to analyze how environmental impact results vary when parameters are adjusted.

- **Monte Carlo Methodology**: The methodology developed can be extended to a plethora of material and product types to gain insights into their environmental impact ranges in specific regional markets.

## Repository Structure

- **comparison data**: Contains EPD datapoints from EC3 for validation of the openimpact model results.
- **models**: Houses packaged JSON-LD files with LCA data compatible with openLCA.
- **providers**: Features tables of provider options and their probabilities.
- **results files**: Contains both raw and processed results, including select results charts.
- **substitutions**: Includes tables of model modifications for Monte Carlo simulations.

## How to Use the Models

1. **Open your database**: Launch your application and open the desired database.
2. **Import the Model**:
   - Navigate to `File` > `Import` > `JSON-LD`.
   - Choose the JSON-LD file you want to import. Ensure it's in the `.zip` format.
     > **Note**: Do not unzip the file. Load it as-is into openLCA.
   - Click `Next` if you wish to set overwrite options.
   - Click `Finish` to complete the import process.
