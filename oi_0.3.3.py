#!/python3.11

import os
import re
import sys
import logging
import shutil
import pytz
import timeit
from datetime import datetime
import humanfriendly
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import random
import uuid
import csv
import olca_ipc as ipc
import olca_schema as olca
from typing import Callable
import subprocess
import time


plt.style.use('ggplot')
plt.ion()

timer_start = timeit.default_timer()

"""
First, enable OpenLCA communication with Python from within the OpenLCA program.
In the top ribbon of OpenLCA open: 'Tools / Developer Tools / IPC Server' and click 'OK'
You may need to redo the above steps anytime you switch databases. Alternatively, you can run this
script as well as launch OLCA IPC server by running "run_oi.py" from terminal.
"""
client = ipc.Client(8080)


def main():
    # # Set up logging
    # log_path = os.path.join("logs", f"logfile_{datetime.today().strftime('%y%m%d-%H%M')}.txt")
    # logging.basicConfig(level=logging.DEBUG, filename=log_path, format="")

    # enable specifying substitution sheet name from command line
    # sub_name = input("Enter filename: ")
    #
    # base_analysis_str = input("Do you want to run base case simulation? (y/n) ")
    # base_analysis = True if base_analysis_str.lower() == 'y' else False
    #
    # range_analysis_str = input("Do you want to run range simulation? (y/n) ")
    # range_analysis = True if range_analysis_str.lower() == 'y' else False
    #
    # probability_analysis_str = input("Do you want to run Monte Carlo simulation? (y/n) ")
    # probability_analysis = True if probability_analysis_str.lower() == 'y' else False
    #
    # if probability_analysis:
    #     # number of provider substitution loops
    #     loop_runs = int(input("How many loop runs (e.g. 60): "))
    #     # number of parameter redefinitions per each loop_run
    #     param_runs = int(input("How many param runs (e.g. 10): "))
    #
    # # Ask for an integer input
    # while True:
    #     try:
    #         max_value = float(input("Enter max GWP for plot (e.g. 5): "))
    #         # logging.info(f"Max value: {max_value}")
    #         # "Suggestions: steel=5 kgCO2e/kg, concrete=700 kgCO2e/m3, "
    #         # "wood=500 kgCO2e/m3, electricity=2 kgCO2e/kWh>")
    #         break
    #     except ValueError:
    #         print("Please enter a valid integer.")

    """
    Select what calculation to run, what substitution sheet to use, and what LCIA method to use.
    This script is setup either for a 'Simple' analysis, or a full 'Monte Carlo' analysis. Simple uses
    standard values defined on OLCA, Monte Carlo uses OLCA-defined uncertainty values.
    MODIFY AS NEEDED: ==========================================================================================
    """
    sub_names = [
        'steel_heavysection_v3',
        'steel_hss_v3',
        'steel_plate_v3',
        'steel_rebar_v3',
        'steel_sheet_galv_v3.2'
    ]

    # 'steel_heavysection_a1a2a3_v2',
    # 'steel_hss_a1a2a3_v2',
    # 'steel_plate_a1a2a3_v2',
    # 'steel_rebar_a1a2a3_v2',
    # 'steel_sheet_galv_a1a2a3_v2'
    # 'steel_sheet_galv_v3.2'

    # Build a product system for use in the calculation? If False, then the main process is used directly
    # without first building a product system. Note that this only works if all default providers are already set.
    calc_using_ps = True

    for sub_name in sub_names:
        # sub_name = 'steel_heavysection_a1a2a3_v2'  # the filename of the substitution sheet
        base_analysis = True  # Do you want to run base case simulation?
        range_analysis = True  # Do you want to run range simulation?
        subgroup_mca = True  # grouped monte carlo simulation aimed at getting uncertainty group variations
        probability_analysis = True  # Do you want to run probabilistic simulation?
        loop_runs = 50
        param_runs = 5
        max_value = 5.0  # kgCO2e/unit, expected highest value for setting plot axis max

        """
        Value suggestions:
        max_value = 5.0  # steel products
        max_value = 700  # concrete or cement products
        max_value = 500  # wood products
        max_value = 2.0  # electricity, per kWh
        """

        lcia_methods = [
            'TRACI 2.1 (openIMPACT)'
            # 'IPCC 2013 GWP 100a',
            # 'EF Method (adapted)',
            # 'CML-IA baseline'
        ]

        """
        END OF MANUAL MODIFICATIONS ====================================================================================
        """

        print(f'Loading "{sub_name}" substitution sheet.')

        """ LOAD SUBSTITUTION DATA """

        # Read and clean substitution sheet
        sub_sheet_path = os.path.join("substitutions", f"{sub_name}.xlsx")
        sub_sheet = pd.read_excel(sub_sheet_path)  # Load the substitution .xlsx file
        sub_sheet = sub_sheet.replace(r'^\s+$', np.nan, regex=True)  # Replace all empty cells with "nan"
        sub_sheet = sub_sheet[~sub_sheet['skip'].isin(['Yes'])]  # Skip any rows marked as skip: Yes

        # List all provider sheets used in this simulation
        print(f'\nIdentifying provider substitution sheets.')
        prov_sheet = sub_sheet[sub_sheet['provider_sheet'].notna()]  # Subset only rows with provider sheets listed
        prov_sheet = prov_sheet.reset_index()
        provider_sheets = prov_sheet['provider_sheet'].dropna().unique().tolist()  # List unique provider sheets

        for i in provider_sheets:
            print(f'\t{i}')

        # List all parameters used in this simulation
        print(f'\nIdentifying parameters and their context.')
        param_sheet = sub_sheet[sub_sheet['parameter'].notna()].reset_index()
        # alternative:
        # mod_filter = sub_sheet.loc[sub_sheet['mod'] == 'parameter']

        param_list = []
        for index, row in param_sheet.iterrows():
            context_name = row['name']
            param_name = row['parameter']
            param_list.append(f'{context_name}.{param_name}')
            print(f'\t{param_list[index]}')

        """ BASE ANALYSIS SET UP """

        """
        Probabilistic analysis can run just with the basic substitution information.
        Base case analysis requires additional data extraction and manipulation below.
        1. Extract the base, low, and high providers from each provider sheet
        2. Extract the base, low, and high parameters for each process from the substitution sheet
        3. Create table of all combinations of providers and parameters
        4. Run OLCA for each combination of inputs
        """

        """ 1. Extract the base, low, and high providers from each provider sheet """
        # List all base, low, and high providers for each provider sheet
        print(f'\nIdentifying base model providers.')
        provider_lists = []  # placeholder list for subset tables

        for sheet_name in provider_sheets:
            try:
                # extract base scenario data from provider sheets
                sheet_path = os.path.join("providers", f"{sheet_name}.xlsx")

                prov_subset = pd.read_excel(sheet_path)  # read the excel file into a pandas dataframe
                prov_subset["provider_sheet"] = sheet_name  # add sheet name to the table
                prov_subset = prov_subset.replace(r'^\s+$', np.nan, regex=True)  # Replace empties and spaces

                # create a filter to only keep rows where the "mark" column contains "base", "low", or "high"
                mark_filter = prov_subset['mark'].isin(['base', 'low', 'high'])
                subset_columns = ['provider_sheet', 'process_uuid', 'location', 'name', 'mark']
                prov_subset = prov_subset.loc[mark_filter, subset_columns]

                # add the subset table to the list of tables
                provider_lists.append(prov_subset)

            except FileNotFoundError:
                print(f'\t!! No such file or directory: {sheet_path} !!\n'
                      f'\t!! Check that {sheet_name} is in the providers folder and matches substitution sheet.')
            except KeyError:
                print(f'\t!! Check errors in provider sheet {sheet_name} !!\n'
                      f'\t!! Check for typos, missing data, etc.')

        # combine all the subset tables into one table
        p_df = pd.concat(provider_lists)

        # create a filter to only keep rows where the "mark" column contains "base"
        base_filter = p_df['mark'] == 'base'
        base_p_df = p_df.loc[base_filter].reset_index()
        range_p_df = p_df.loc[~base_filter].reset_index()

        print(f'Base model providers loaded.')

        # print(f'\nCheck base providers:')
        # for index, row in base_p_df.iterrows():
        #     print(f"{row['provider_sheet']}:\t {row['name']}")
        #
        # print(f'\nCheck range providers:')
        # for index, row in range_p_df.iterrows():
        #     print(f"{row['provider_sheet']}.{row['mark']}:\t {row['name']}")

        """ 2. Extract the base, low, and high parameters for each process from the substitution sheet """
        # List all base, low, and high parameters for each provider sheet
        print(f'\nIdentifying base model parameters.')

        param_lists = []  # placeholder list for param tables
        for index, row in param_sheet.iterrows():
            q_base = pick_value(row['sample'], "base")
            q_low = pick_value(row['sample'], "low")
            q_high = pick_value(row['sample'], "high")
            param_lists.append([row['uuid'], row['name'], row['parameter'], "base", q_base])
            param_lists.append([row['uuid'], row['name'], row['parameter'], "low", q_low])
            param_lists.append([row['uuid'], row['name'], row['parameter'], "high", q_high])

            # check parameter selections
            # print(f'{index+1}) {row["name"]}.{row["parameter"]}\n'
            #       f'\t{row["sample"]}\n'
            #       f'\t\tBase value: {q_base}\n'
            #       f'\t\tLow value: {q_low}\n'
            #       f'\t\tHigh value: {q_high}\n')

        # select the columns you want to include in the subset table
        param_columns = ['uuid', 'name', 'parameter', 'mark', 'value']
        q_df = pd.DataFrame(param_lists, columns=param_columns)

        # create a filter to only keep rows where the "mark" column contains "base"
        base_filter = q_df['mark'] == 'base'
        base_q_df = q_df.loc[base_filter].reset_index()
        range_q_df = q_df.loc[~base_filter].reset_index()

        print(f'Base model parameters loaded.')

        # print(f'\nCheck base parameters:')
        # for index, row in base_q_df.iterrows():
        #     print(f"{row['name']}.{row['parameter']}: {row['value']}")
        #
        # print(f'\nCheck range parameters:')
        # for index, row in range_q_df.iterrows():
        #     print(f"{row['name']}.{row['parameter']}.{row['mark']}: {row['value']}")

        """ GENERAL SET UP """

        """
        Test for splitting find_flow and regions columns in the substitution sheet. Note that regional selection is
        currently hardcoded and the sheet's region column is not being used. The idea is to enable region specification
        from the substitution sheet in the future.
        """
        # print(prov_sheet['find_flow'].iloc[1].replace('", "', ';').replace('"', '').split(';'))
        # print(prov_sheet['regions'].iloc[2].replace('", "', ';').replace('"', '').split(';'))

        # Get the reference amount and unit of the main process (from which a Product System is later created).
        main_process_uuid = sub_sheet['uuid'].iloc[0]  # Get main process uuid
        main_process_json = fetch_process_json(main_process_uuid)  # Fetch main process JSON
        ref_amount, ref_unit = find_ref_flow(main_process_json)  # Get main process reference flow info
        print(f'\nGetting product system reference info.\n'
              f'\tMain process:\t {main_process_json.name}\n'
              f'\tAmount:\t\t\t {ref_amount}\n'
              f'\tUnit:\t\t\t {ref_unit}')

        # Setup results directory.
        # Create a results directory if it doesn't already exist.
        my_dir = os.path.join("results files", f"{main_process_json.name}", "raw")
        if not os.path.exists(my_dir):
            os.makedirs(my_dir)

        """ SETUP OF RESULTS FILES """

        """ for MCA """
        print(f'\nSetting up a csv results file for Probabilistic Analysis.')
        datetime_stamp = datetime.today().strftime('%y%m%d-%H%M')
        results_name = f'{main_process_json.name} {datetime_stamp}'
        res_path = os.path.join("results files", f"{main_process_json.name}", "raw", f"{results_name}.csv")
        header = (provider_sheets + param_list +
                  ['gwp', 'gwp_be', 'gep_bu', 'ap', 'ep', 'odp', 'pocp', 'gwp_AR5', 'gwp_EF2', 'gwp_CML'] +
                  ["sim_type"])

        f = open(res_path, "w", newline='')
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        f.close()

        # Create a directory with substitution and provider files for debugging
        debug_dir = os.path.join("results files", f"{main_process_json.name}", "raw", f"{results_name} input debug")
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)

        # Copy each xlsx file to the results directory
        source_subs = os.path.join("substitutions", f"{sub_name}.xlsx")
        debug_subs = os.path.join(debug_dir, f"{sub_name}.xlsx")
        shutil.copyfile(source_subs, debug_subs)
        for sheet_name in provider_sheets:
            source_providers = os.path.join("providers", f"{sheet_name}.xlsx")
            debug_providers = os.path.join(debug_dir, f"{sheet_name}.xlsx")
            shutil.copyfile(source_providers, debug_providers)

        """ BASE SIMULATION """

        """
        1. Run base simulation
            Set all providers to base providers
            Set all parameters to base parameters
            Run simulation
    
        2. For each row in range_p_df (range provider table)
            Set all providers to base providers
            Sub range provider
            Set all parameters to base parameters
            Run simulation
    
        3. For each row in range_q_df (range parameter table)
            Set all providers to base providers
            Set all parameters to base parameters
            Sub range parameter
            Run simulation
        """

        if base_analysis:
            print(f'\nStarting base simulation.\n=============================================================')
            base_start = timeit.default_timer()

            counter = 0

            """ 1. Run base simulation """

            """SETUP: PROVIDERS OF FLOWS"""
            print(f'\nGetting base provider data')
            provider_dict = identify_providers(base_p_df)

            print(f'\nModifying all relevant processes.')
            modify_processes(prov_sheet, provider_dict)

            if calc_using_ps:
                """SETUP: PRODUCT SYSTEM"""
                print(f'\nCreating a product system.')
                model_ref = create_ps(main_process_json)

            """SETUP: PARAMETER REDEFINITION"""
            print(f'\nGetting base parameter data')
            param_picked = []
            parameter_redefs = []
            for index, row in base_q_df.iterrows():
                # Append picked value to list of redefinitions for results sheet
                param_picked.append(row['value'])
                # Redefine parameters in OLCA model
                redef = olca.ParameterRedef(
                    context=client.get_descriptor(olca.Process, row['uuid']),
                    name=row['parameter'],
                    value=row['value']
                )
                parameter_redefs.append(redef)
                # print(f"\t{index}) {row['parameter']}: {row['value']}")
                # print(f"\t{q_index}) {q_row['parameter']}: {value} || {redef}")

            """EXECUTE: CALCULATION"""
            counter += 1
            impact_results = []
            impact_results = get_results(model_ref, lcia_methods, counter, parameter_redefs)

            """
            Save results to csv. A new csv file is created on every first simulation and any additional runs in 
            that simulation are appended to that csv. This ensures that results are saved even if the full 
            simulation doesn't finish.
            """
            # provider_keys_list = list(provider_dict)
            providers_picked = []
            for sheet in provider_sheets:
                providers_picked.append(provider_dict[sheet].name)

            fields = providers_picked + param_picked + impact_results + ["base"]

            # Append results to an existing csv file
            with open(res_path, 'a', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(fields)
            # counter += 1
            # print(f'\nResult saved to csv. | Run {counter} gwp: {gwp:.2f} {gwp_unit} ({lcia_methods[0]})')

            if calc_using_ps:
                client.delete(model_ref)  # delete product system

            print('\nElapsed time: {}'.format(humanfriendly.format_timespan(timeit.default_timer() - timer_start)))

            """ 2. For each row in range_p_df (range provider table) """

        if range_analysis:
            print(f'\n\nStarting provider range simulations...')
            # for all providers run through swaps between base and range providers
            for range_index, range_row in range_p_df.iterrows():

                print(f'\nResetting all providers to base selection')
                # identifying all base providers (base_p)
                provider_dict = identify_providers(base_p_df)

                print(f"\nSetting {range_row['provider_sheet']} to {range_row['mark']}")
                # If UUID is provided, get REF by UUID, else get REF by Name.
                if isinstance(range_row['process_uuid'], str):
                    provider_ref = client.get_descriptor(olca.Process, range_row['process_uuid'])
                else:
                    provider_ref = client.find(olca.Process, range_row['name'])

                # save the range provider selection to provider_dict for this iteration
                provider_dict[range_row['provider_sheet']] = provider_ref

                print(f'\nModifying all relevant processes.')
                # modify all processes using either the selected base or range provider listed in provider_dict
                modify_processes(prov_sheet, provider_dict)

                if calc_using_ps:
                    print(f'Creating a product system.')
                    model_ref = create_ps(main_process_json)

                print(f'\nGetting base parameter data')
                # set_params(base_q_df)

                parameter_dict = {}
                param_picked = []
                parameter_redefs = []

                # set all parameters to base (base_q)
                for index, row in base_q_df.iterrows():
                    # Append picked value to list of redefinitions for results sheet
                    param_picked.append(row['value'])
                    # Redefine parameters in OLCA model
                    redef = olca.ParameterRedef(
                        context=client.get_descriptor(olca.Process, row['uuid']),
                        name=row['parameter'],
                        value=row['value']
                    )
                    parameter_redefs.append(redef)
                    print(f"\t{index}) {row['parameter']} :: {row['value']}")
                    # print(f"\t{q_index}) {q_row['parameter']}: {value} || {redef}")

                """EXECUTE: CALCULATION"""

                counter += 1
                impact_results = []
                impact_results = get_results(model_ref, lcia_methods, counter, parameter_redefs)

                """
                Save results to csv. A new csv file is created on every first simulation and any additional runs in 
                that simulation are appended to that csv. This ensures that results are saved even if the full 
                simulation doesn't finish.
                """
                # provider_keys_list = list(provider_dict)
                providers_picked = []
                for sheet in provider_sheets:
                    providers_picked.append(provider_dict[sheet].name)

                fields = providers_picked + param_picked + impact_results + ["range"]

                with open(res_path, 'a', newline='') as f:  # 'a' appends to an existing file
                    writer = csv.writer(f)
                    writer.writerow(fields)
                # counter += 1
                # print(f'\nResult saved to csv. | Run {counter} gwp: {gwp:.2f} {gwp_unit} ({lcia_methods[0]})')

                if calc_using_ps:
                    client.delete(model_ref)  # delete product system

                print('\nElapsed time: {}'.format(humanfriendly.format_timespan(timeit.default_timer() - timer_start)))

            """ 3. For each row in range_q_df (range parameter table) """

            print(f'\n\nStarting parameter range simulations...')
            # for all parameters run through swaps between base and range parameters
            for range_index, range_row in range_q_df.iterrows():

                print(f'\nResetting all providers to base selection')
                provider_dict = identify_providers(base_p_df)

                print(f'\nModifying all relevant processes.')
                modify_processes(prov_sheet, provider_dict)

                if calc_using_ps:
                    print(f'Creating a product system.')
                    model_ref = create_ps(main_process_json)

                print(f'\nGetting base parameter data')

                param_picked = []
                parameter_redefs = []

                for base_index, base_row in base_q_df.iterrows():
                    if base_row['name'] == range_row['name'] and base_row['parameter'] == range_row['parameter']:
                        # Append picked value to list of redefinitions for results sheet
                        param_picked.append(range_row['value'])
                        # Redefine parameters in OLCA model
                        redef = olca.ParameterRedef(
                            context=client.get_descriptor(olca.Process, range_row['uuid']),
                            name=range_row['parameter'],
                            value=range_row['value']
                        )
                        parameter_redefs.append(redef)
                        print(f"\t{range_index}) {range_row['mark']} value of {range_row['value']} "
                              f"set for {range_row['name']}.{range_row['parameter']}")
                        # print(f"\t{q_index}) {q_row['parameter']}: {value} || {redef}")

                    else:
                        # Append picked value to list of redefinitions for results sheet
                        param_picked.append(base_row['value'])
                        # Redefine parameters in OLCA model
                        redef = olca.ParameterRedef(
                            context=client.get_descriptor(olca.Process, base_row['uuid']),
                            name=base_row['parameter'],
                            value=base_row['value']
                        )
                        parameter_redefs.append(redef)
                        print(f"\t{base_index}) {base_row['mark']} value of {base_row['value']} "
                              f"set for {base_row['name']}.{base_row['parameter']}")
                        # print(f"\t{q_index}) {q_row['parameter']}: {value} || {redef}")

                """EXECUTE: CALCULATION"""

                counter += 1
                impact_results = []
                impact_results = get_results(model_ref, lcia_methods, counter, parameter_redefs)

                """
                Save results to csv. A new csv file is created on every first simulation and any additional runs in 
                that simulation are appended to that csv. This ensures that results are saved even if the full 
                simulation doesn't finish.
                """
                # provider_keys_list = list(provider_dict)
                providers_picked = []
                for sheet in provider_sheets:
                    providers_picked.append(provider_dict[sheet].name)

                fields = providers_picked + param_picked + impact_results + ["range"]

                with open(res_path, 'a', newline='') as f:  # 'a' appends to an existing file
                    writer = csv.writer(f)
                    writer.writerow(fields)
                # counter += 1
                # print(f'\nResult saved to csv. | Run {counter} gwp: {gwp:.2f} {gwp_unit} ({lcia_methods[0]})')

                if calc_using_ps:
                    client.delete(model_ref)  # delete product system

                print('\nElapsed time: {}'.format(humanfriendly.format_timespan(timeit.default_timer() - timer_start)))

            """ End of base simulation """
            # Show execution time for base simulation.
            print(f'\nTotal base scenario run time: {humanfriendly.format_timespan(timeit.default_timer() - base_start)}')

        """ MONTE CARLO SIMULATION """

        """
        1. For as many iterations as specified
            1. Pick provider from each provider sheet by probability
            2. Modify all providers across the entire system
            3. Create OLCA product system object
            4. Create OLCA calculation set up
            5. Set parameters for the calculation
            6. Run calculations for each LCIA method given the product system and parameters
            7. Save results for this run and continue with step 1.1 or 1.5
                - it is simpler/quicker to redefine parameters, that's why do multiple loops of just parameter re-def's
                - it is more complicated and takes longer to swap providers
        """

        if probability_analysis:
            print(f'\nStarting probability simulation.\n=============================================================')
            mca_start = timeit.default_timer()

            """
            Specify number of loop runs and OLCA simulator runs. Loop runs randomly sample from a set of providers, modify
            and update all of the main OLCA processes, update the final ProductSystem and proceed to the OLCA simulation.
            OLCA simulations consist of OLCA sampling from the basic and DQI uncertainty definitions associated with
            processes, exchanges, and/or parameters.
            """
            counter = 0

            for run in range(loop_runs):
                loop_timer_start = timeit.default_timer()
                print(f"\n\nStarting iteration {run+1} / {loop_runs} =======================================================")
                """
                Load picked provider JSONs into a dictionary which can later be accessed without re-sampling and re-calling
                the providers for the same monte carlo run. This also ensures only one provider is picked for the same
                type of provider substitution, e.g. that all substitutions of electricity flows across all processes source
                PJM Interconnection grid electricity, and not multiple varying electricity providers.
                """
                print(f'\nPicking providers')
                provider_dict = {}
                for sheet_name in provider_sheets:
                    try:
                        """
                        Get the full path to the current provider sheet and run the pick provider function.
                        If needed, provide a list of regions in the sample_provider function call to filter specific regions.
                        """
                        sheet_path = f'./providers/{sheet_name}.xlsx'
                        provider_uuid, provider_name, provider_location = sample_provider(sheet_path)

                        # If UUID is provided, get REF by UUID, else get REF by Name.
                        if isinstance(provider_uuid, str):
                            provider_ref = client.get_descriptor(olca.Process, provider_uuid)
                        else:
                            provider_ref = client.find(olca.Process, provider_name)

                        provider_dict[sheet_name] = provider_ref  # Save to provider_dict
                        print(f'\tFrom "{sheet_name}" picked "{provider_ref.name}" and loaded json.')
                    except FileNotFoundError:
                        print(f'\t!! No such file or directory: {sheet_path} !! MCA will terminate !!')
                    except KeyError:
                        print(f'!! Check errors in provider sheet {sheet_path} !! MCA will terminate !!')

                """Displaying provider dict for QA purposes. Not critical."""
                # print(f'\nprovider_dict')
                # for key in provider_dict:
                #     print(f'\n\tkey:\t\t {key}\n'
                #           f'\tvalue.id:\t {provider_dict[key].id}\n'
                #           f'\tvalue.name:\t {provider_dict[key].name}')

                """
                Modify all target processes with the new providers.
                """

                print(f'\nModifying all relevant processes.')
                modify_processes(prov_sheet, provider_dict)

                """
                Create new temporary product system for the main process. This step takes into account all provider 
                substitutions made in previous steps. A new product system has to be created anytime there is a change
                to any Process because the OLCA Update function does not work for this. The product system is deleted
                at the end of each loop so it does not clutter the database.
                """
                if calc_using_ps:
                    print(f'Creating a product system.')
                    model_ref = create_ps(main_process_json)

                """
                Product system parameter definitions setup.
                Redefining the product system parameters, and then setting up the product system for this calculation.
                """
                for param_loop in range(param_runs):
                    print(f"\nPicking parameters. Parameter redefinition loop {param_loop+1} / {param_runs} ----\n")

                    param_picked = []
                    parameter_redefs = []

                    for q_index, q_row in param_sheet.iterrows():
                        context_name = q_row['name']
                        param_name = q_row['parameter']

                        param_string = q_row['sample']  # Take the sample from substitution sheet
                        value = pick_value(param_string, "sample")  # Pick parameter value based on sample information
                        # Append picked value to list of redefinitions for results sheet
                        param_picked.append(value)
                        # Redefine parameters in OLCA model
                        redef = olca.ParameterRedef(
                            context=client.get_descriptor(olca.Process, q_row['uuid']),
                            name=q_row['parameter'],
                            value=value
                        )

                        # print(f'{context_name} {param_name}\n'
                        #       f'\t{param_string}\n'
                        #       f'\tPicked value: {value}\n')

                        parameter_redefs.append(redef)
                        print(f"\t{q_index}) {q_row['parameter']} :: {value}")
                        # print(f"\t{q_index}) {q_row['parameter']}: {value} || {redef}")

                    """ RUN SIMULATION"""

                    counter += 1
                    impact_results = []
                    impact_results = get_results(model_ref, lcia_methods, counter, parameter_redefs)

                    """
                    Save results to csv. A new csv file is created on every first simulation and any additional runs in 
                    that simulation are appended to that csv. This ensures that results are saved even if the full 
                    simulation doesn't finish.
                    """
                    # provider_keys_list = list(provider_dict)
                    providers_picked = []
                    for sheet in provider_sheets:
                        providers_picked.append(provider_dict[sheet].name)

                    fields = providers_picked + param_picked + impact_results + ["mca"]

                    with open(res_path, 'a', newline='') as f:  # 'a' appends to an existing file
                        writer = csv.writer(f)
                        writer.writerow(fields)
                    # counter += 1
                    # print(f'\nResult saved to csv. | Run {counter} gwp: {gwp:.2f} {gwp_unit} ({lcia_methods[0]})')

                    """
                    Plot latest histogram. If you do not wish to view live histogram updates, then comment this out.
                    """
                    display_result(main_process_json.name,
                                   declared_unit=ref_unit,
                                   results_path=res_path,
                                   max_value=max_value)

                """
                Delete product system so that we can create a new one in the next loop without cluttering the database.
                This is currently needed because the product system update functionality in openLCA does not work,
                which means that even after modifying upstream processes running the simulation using the same
                product system would have yielded identical results as before process modifications.
                """
                if calc_using_ps:
                    client.delete(model_ref)  # Delete product system

                time_left = (timeit.default_timer() - loop_timer_start)*(loop_runs - run)
                print('\nElapsed time: {}'.format(humanfriendly.format_timespan(timeit.default_timer() - timer_start)))
                print('Time remaining: {}'.format(humanfriendly.format_timespan(time_left)))

            # Show execution time for probability simulation.
            print(f'\nTotal MCA run time: {humanfriendly.format_timespan(timeit.default_timer() - mca_start)}')

        if subgroup_mca:
            print(f'\nStarting sub-group simulation.\n=============================================================')
            mca_start = timeit.default_timer()
            counter = 0

            # get unique uf_groups
            uf_groups = sub_sheet['uf_group'].dropna().unique()
            print(uf_groups)

            for ufg in uf_groups:
                print(f'\nRunning MCA for "{ufg}" variation')

                # set all to base providers
                print(f'\nGetting base provider data')
                provider_dict = identify_providers(base_p_df)

                # list provider sheets and parameters in this group
                group_p_df = prov_sheet[prov_sheet['uf_group'] == ufg]  # group provider dataframe
                group_q_df = param_sheet[param_sheet['uf_group'] == ufg]  # group parameter dataframe

                for run in range(loop_runs):
                    loop_timer_start = timeit.default_timer()
                    print(f"\n\nStarting iteration {run + 1} / {loop_runs} ==============================================")

                    print(f'\nPicking providers')
                    for index, row in group_p_df.iterrows():
                        try:
                            sheet_name = row['provider_sheet']
                            sheet_path = f'./providers/{sheet_name}.xlsx'
                            provider_uuid, provider_name, provider_location = sample_provider(sheet_path)

                            # If UUID is provided, get REF by UUID, else get REF by Name.
                            if isinstance(provider_uuid, str):
                                provider_ref = client.get_descriptor(olca.Process, provider_uuid)
                            else:
                                provider_ref = client.find(olca.Process, provider_name)

                            provider_dict[sheet_name] = provider_ref  # Save to provider_dict
                            print(f'\t"{sheet_name}" :: "{provider_ref.name}"')
                        except FileNotFoundError:
                            print(f'\t!! No such file or directory: {sheet_path} !! MCA will terminate !!')
                        except KeyError:
                            print(f'!! Check errors in provider sheet {sheet_path} !! MCA will terminate !!')

                    print(f'\nModifying all relevant processes.')
                    modify_processes(prov_sheet, provider_dict)

                    if calc_using_ps:
                        print(f'Creating a product system.')
                        model_ref = create_ps(main_process_json)

                    # redefine all parameters as necessary
                    for param_loop in range(param_runs):
                        print(f"\nPicking parameters. Parameter redefinition loop {param_loop+1} / {param_runs} ----\n")

                        print(f'\nSetting base parameter data')
                        param_picked = []
                        parameter_redefs = []

                        # if in ufg group then sample randomly, else assign base parameter
                        for q_index, q_row in param_sheet.iterrows():
                            if q_row['uf_group'] == ufg:
                                # Pick parameter value based on sample information
                                q_value = pick_value(q_row['sample'], "sample")
                                print(f"\t\t{q_index}) "
                                      f"{q_row['name']}.{q_row['parameter']} :: {q_value} :: random :: {ufg}")
                            else:
                                # Pick base value
                                q_value = pick_value(q_row['sample'], "base")
                                print(f"\t\t{q_index}) "
                                      f"{q_row['name']}.{q_row['parameter']} :: {q_value} :: base")

                            # Append picked value to list of redefinitions for results sheet
                            param_picked.append(q_value)

                            # Redefine parameters in OLCA model
                            redef = olca.ParameterRedef(
                                context=client.get_descriptor(olca.Process, q_row['uuid']),
                                name=q_row['parameter'],
                                value=q_value
                            )
                            parameter_redefs.append(redef)

                        counter += 1
                        impact_results = []
                        impact_results = get_results(model_ref, lcia_methods, counter, parameter_redefs)

                        # provider_keys_list = list(provider_dict)
                        providers_picked = []
                        for sheet in provider_sheets:
                            providers_picked.append(provider_dict[sheet].name)

                        fields = providers_picked + param_picked + impact_results + [ufg]

                        # appends results to an existing csv file
                        with open(res_path, 'a', newline='') as f:
                            writer = csv.writer(f)
                            writer.writerow(fields)

                        display_result(main_process_json.name,
                                       declared_unit=ref_unit,
                                       results_path=res_path,
                                       max_value=max_value)

                    if calc_using_ps:
                        client.delete(model_ref)  # Delete product system

        # Show elapsed execution time.
        print('\nTotal run time: {}'.format(humanfriendly.format_timespan(timeit.default_timer() - timer_start)))


"""END OF MAIN"""


"""OPEN LCA MANIPULATION FUNCTIONS"""


def identify_providers(provider_df):
    provider_dict = {}
    for index, row in provider_df.iterrows():
        sheet_name = row['provider_sheet']
        provider_uuid = row['process_uuid']
        provider_name = row['name']
        provider_location = row['location']

        # If UUID is provided, get REF by UUID, else get REF by Name.
        if isinstance(provider_uuid, str):
            provider_ref = client.get_descriptor(olca.Process, provider_uuid)  # Get JSON for the selected provider
        else:
            provider_ref = client.find(olca.Process, provider_name)  # Get REF of selected provider

        provider_dict[sheet_name] = provider_ref  # Save to provider_dict

        print(f"\t{sheet_name} :: {provider_ref.name}")

    return provider_dict


def modify_processes(prov_sheet, provider_dict):
    modify_start = timeit.default_timer()

    prov_sheet_rows = prov_sheet.shape[0]  # Total number of provider sheets to run through

    for index, row in prov_sheet.iterrows():
        print(f'\n{index + 1} / {prov_sheet_rows}')
        try:
            process_json = fetch_process_json(row['uuid'])
            # print(f'\tProcess to mod: {process_json.name}')
            find_flows = row['find_flow'].replace('", "', '++').replace('"', '').split('++')
            # print(f'\tFlows to mod: {find_flows}')
            provider_ref = provider_dict[row['provider_sheet']]
            # print(f'\tProvider name: {provider_json.name}\n')
            modified = True

        except KeyError:
            process_json = client.get_descriptor(olca.Process, [row['uuid']])
            print(f'!! No changes made to {process_json.name} due to KeyError.\n')
            modified = False
        except AttributeError:
            process_json = client.get_descriptor(olca.Process, [row['uuid']])
            print(f'!! No changes made to {process_json.name} due to AttributeError.\n')
            modified = False

        if modified:
            modify_exchanges(
                process=process_json,
                find_flow=find_flows,
                new_provider=provider_ref
            )

    modify_time = humanfriendly.format_timespan(timeit.default_timer() - modify_start)
    print(f'\n\nModifications finished in {modify_time}.')


def pick_value(param_string, mark="base"):
    """
    Picks a sample value based on the specified distribution and parameters from a substitution sheet.
    E.g., "triangular; min=0.01; mode=0.0771; max=0.08".

    :param param_string: string from substitution sheet column "sample"
    :param mark: value type from options of: "sample", "base", "high", "low"
        sample: uses the provided string to sample from a probability distribution
        base: base value, usually mean or median
        high: value that is at the higher end of impact
        low: value that is at the lower end of impact
    :return: selected value
    """

    pars = param_string.split(';')
    # print(f'{pars}')

    if pars[0] == "list":
        pars_list = pars[1].split(",")

        sample = float(np.random.choice(pars_list))
        base = float(pars[2].split("=", 1)[1])
        high = max(pars_list)
        low = min(pars_list)

    else:
        for i in range(1, len(pars)):
            try:
                pars[i] = float(pars[i].split("=", 1)[1])
            except AttributeError:
                pass

        if pars[0] == "uniform":
            sample = np.random.uniform(pars[1], pars[2])  # min, max, size
            base = pars[3]
            high = pars[1]
            low = pars[2]
        if pars[0] == "triangular":
            sample = np.random.triangular(pars[1], pars[2], pars[3])  # left, mode, right
            base = pars[4]
            high = pars[3]
            low = pars[1]
        if pars[0] == "normal":
            sample = np.random.normal(pars[1], pars[2])  # mean, stdv, size
            base = pars[3]
            high = pars[1] + pars[2]
            low = pars[1] - pars[2]
        if pars[0] == "lognormal":
            print(f"Warning! Lognormal sampling is not configured yet!")
            sample = np.random.lognormal(pars[1], pars[2])  # mean, sigma, size
            base = pars[3]
            high = pars[1] + pars[2]
            low = pars[1] - pars[2]

    # print(pars)

    if mark == "sample":
        value = sample
    elif mark == "base":
        value = base
    elif mark == "high":
        value = high
    elif mark == "low":
        value = low
    else:
        print("Invalid sample. Check substitution sheet column: sample.")

    return value


def sample_provider(path, regions=None):
    """
    Loads market share spreadsheet, calculates probability from total amount produced by each provider,
    and randomly selects a single provider using the underlying probability.

    :param path: Excel sheet path
    :param regions: which countries to include?
    :return: single provider based on probability
    """

    """Load provider sheet and remove unnecessary rows."""
    prod_stats = pd.read_excel(path).dropna(subset=["amount"])
    prod_stats = prod_stats.replace(r'^\s+$', np.nan, regex=True)  # Replace empties and spaces with "nan"
    prod_stats = prod_stats[~prod_stats['skip'].isin(['Yes'])]  # Skip any marked rows
    prod_stats = prod_stats.reset_index()  # Reindex the dataframe for consistency in calling by index

    """If specific regions are listed, filter out just those regions, else include everything."""
    if regions is not None:
        prod_stats = prod_stats[prod_stats['region'].isin(regions)]  # Select subset of regions
        print(f'\t\t\tSelecting only from the following regions: {regions}')

    """Recalculate probability based on the amount column for the remaining subset of data."""
    total_prod = prod_stats['amount'].sum()
    prod_stats["share"] = prod_stats["amount"] / total_prod

    """Check that the recalculated sum equals 100% or notify the user of a possible error."""
    if 0.96 < prod_stats['share'].sum() < 1.04:
        pass
    else:
        print(f"\t\t,- Market share sum check failed. Please check market share data and code.\n"
              f"\t\t|- sum({prod_stats['share'].tolist()}) = {prod_stats['share'].sum()}")

    """Prep data and make a random choice selection."""
    names = prod_stats.index.values.tolist()
    shares = prod_stats['share'].tolist()
    selection = np.random.choice(names, p=shares)

    """Save info for the randomly selected process."""
    process_uuid = prod_stats['process_uuid'].iloc[selection]
    process_name = prod_stats['name'].iloc[selection]
    process_location = prod_stats['location'].iloc[selection]

    # print(f'{selection}')
    # print(f"\t\t\t{process_uuid} | {process_name} - {process_location}")

    return process_uuid, process_name, process_location


def modify_exchanges_1(process, find_flow, sub_flow, sub_provider):
    """
    Finds specific exchanges in a process and modifies their flow and default provider.
    :param process: Process whose exchanges will be modified.
    :param find_flow: List of flow names that could represent the exchange that needs modifying.
    :param sub_flow: The flow to be substituted into the process.
    :param sub_provider: The default provider to be substituted into the process.
    :return: None. The process is simply updated in OLCA.
    """
    exchange_list = []
    for i in process.exchanges:
        if i.is_input and i.flow.name in find_flow:
            i.flow = client.find(olca.Flow, sub_flow)
            i.default_provider = sub_provider
        exchange_list.append(i)

    process.exchanges = exchange_list
    process.olca_type = 'Process'

    client.put(process)


def modify_exchanges(process, find_flow, new_provider, preloaded_provider_dict="None"):
    """
    Finds natural gas exchanges in a process and modifies their flow and default provider as well as converts
    units if needed.
    :param process: Process whose exchanges will be modified. Expects OLCA JSON or REF.
    :param find_flow: List of flow names that could represent the exchange that needs modifying.
    :param new_provider: The default provider to be substituted into the process. Expects OLCA JSON or REF.
    :param preloaded_provider_dict: Dictionary with all provider OLCA references, including reference flows. Optional,
    but speeds things up when provided because reference flows do not need to be searched in OLCA.
    :return: None. The process is simply updated in OLCA.
    """
    proc2_mod_json = fetch_process_json(process.id)
    print(f'\tModifying "{proc2_mod_json.name}"')

    proc2_link_json = client.get(olca.Process, new_provider.id)
    proc2_link_ref = client.get_descriptor(olca.Process, new_provider.id)
    print(f'\t\tLinking "{proc2_link_json.name}"')

    if preloaded_provider_dict == "None":
        for i in proc2_link_json.exchanges:
            if i.is_quantitative_reference:
                flow2_link_name = i.flow.name
                flow2_link_uuid = i.flow.id
                flow2_link_type = i.flow_property.name
                flow2_link_unit = i.unit.name
                flow2_link_ref = fetch_flow(flow2_link_name)
    else:
        flow2_link_name = preloaded_provider_dict[proc2_link_json.name]['FlowName']
        flow2_link_uuid = preloaded_provider_dict[proc2_link_json.name]["FlowUuid"]
        flow2_link_type = preloaded_provider_dict[proc2_link_json.name]["FlowType"]
        flow2_link_unit = preloaded_provider_dict[proc2_link_json.name]["FlowUnit"]
        flow2_link_ref = preloaded_provider_dict[proc2_link_json.name]["FlowRef"]

    exchange_list = []
    modifications = 0

    for i in proc2_mod_json.exchanges:
        if i.is_input and i.flow.name in find_flow:
            print(f'\t\t\tTo flow "{i.flow.name}"')
            i.flow = flow2_link_ref
            i.default_provider = proc2_link_ref
            print(f'\t\t\t\tOld amount: {i.amount} {i.unit.name}')

            """
            Check if the flow type matches. If it does not then it needs to be converted. The following conversions
            are supported at this point:
              natural gas: m3 --> MJ
              natural gas: MJ --> m3
            """
            flow2_mod_type = 0
            flow2_mod_unit = 0

            try:
                flow2_mod_type = i.flow_property.name
                flow2_mod_unit = i.unit.name
            except AttributeError:
                print(f'\t\t\t\tNo flow property. Checking direct unit match (e.g., MJ -> MJ)')
                flow2_mod_unit = i.unit.name

            if flow2_mod_type == flow2_link_type:
                pass
            elif flow2_mod_unit == flow2_link_unit:
                pass
            elif 'natural gas' in i.flow.name and i.unit.name == 'm3' and flow2_link_type == 'Energy':
                i.unit = client.find(olca.Unit, 'MJ')
                i.amount = i.amount * 38  # 38 MJ/m3
                try:
                    i.uncertainty.geom_mean = i.uncertainty.geom_mean * 38
                except TypeError:
                    print(f"\t\t\t\tUncertainty conversion failed. If you are using OpenLCA's native uncertainty"
                          f"factors your  results may have issues. If you are running a 'Simple Analysis' with "
                          f"uncertainty handled via OpenIMPACT script then your results will not be affected by "
                          f"this error.")
                except AttributeError:
                    print(f"\t\t\t\tUncertainty conversion failed. It seems there is no uncertainty data in the "
                          f"OpenLCA model.")
            elif 'natural gas' in i.flow.name and i.unit.name == 'MJ' and flow2_link_type == 'Volume':
                i.unit = client.find(olca.Unit, 'm3')
                i.amount = i.amount / 38  # 38 MJ/m3
                try:
                    i.uncertainty.geom_mean = i.uncertainty.geom_mean / 38
                except TypeError:
                    print(f"\t\t\t\tUncertainty conversion failed. If you are using OpenLCA's native uncertainty"
                          f"factors your  results may have issues. If you are running a 'Simple Analysis' with "
                          f"uncertainty handled via OpenIMPACT script then your results will not be affected by "
                          f"this error.")
                except AttributeError:
                    print(f"\t\t\t\tUncertainty conversion failed. It seems there is no uncertainty data in the "
                          f"OpenLCA model.")
            else:
                print(f'\t\t\t\tUnit mismatch during modification! Potential for wrong results.')

            print(f'\t\t\t\tNew amount: {i.amount} {i.unit.name}')
            modifications += 1
        else:
            pass

        exchange_list.append(i)

    proc2_mod_json.exchanges = exchange_list
    proc2_mod_json.olca_type = 'Process'

    client.put(proc2_mod_json)

    if modifications == 0:
        print(f"\t\t!! {find_flow}\n\t\t !! Not modified. Check your process names and uuids in the substitution"
              f"\n\t\t!! sheet and make sure they match names and uuids exactly as they appear in openLCA.")
    else:
        print(f"\t\tCompleted {modifications} modification(s).")

    return proc2_mod_json, modifications


def find_ref_flow(process):
    for i in process.exchanges:
        if i.is_quantitative_reference:
            ref_amount = i.amount
            ref_unit = i.unit.name
            ref_uuid = i.flow.id

    return ref_amount, ref_unit


def create_ps(process_ref):
    """
    Creates new product system. This is necessary to map the whole supply chain after any provider substitution.
    :param process_ref: Expects OLCA reference to top-level process that is to be made into a product system.
    :return: New product system
    """

    # Create a new product system from an existing process. Note that this creates a new Ref object only.
    # You may need to close and reopen the working database in OpenLCA to see the new product system in OpenLCA.
    print(f"Working on {process_ref.name[:50]}. This may take a minute, please wait.")

    # Configure product system creation
    config = olca.LinkingConfig(
        prefer_unit_processes=True,
        provider_linking=olca.ProviderLinking.PREFER_DEFAULTS,
    )

    # Create product system
    new_ps_ref = client.create_product_system(process_ref, config)

    # Next, we need to fetch the whole JSON object and modify its attributes.
    new_ps = client.get(olca.ProductSystem, new_ps_ref.id)
    new_ps.olca_type = olca.schema.ProductSystem.__name__
    new_ps.name = f"{process_ref.name}"
    new_ps.version = "0.0.1"
    new_ps.last_change = datetime.now(pytz.utc).isoformat()  # log current date and time
    new_ps.description = (
        f"This is a product system created using olca ipc and the BT MCA script."
        f"Linking approach during creation: Only default providers; Preferred process type: Unit process"
    )

    print(f"Updating product system. This may take a minute, please wait.")
    update_start = timeit.default_timer()  # start timer
    client.put(new_ps)  # update product system in OLCA
    update_time = humanfriendly.format_timespan(timeit.default_timer() - update_start)  # stop timer
    print(f'Update finished in {update_time}.')

    return new_ps


def display_result(study_object, declared_unit, results_path, max_value):
    """
    Set up result display.
    """
    data = pd.read_csv(f'{results_path}')
    hist_gwp_vals = data['gwp'].tolist()

    plt.hist(hist_gwp_vals, density=True, bins=50, range=[0, max_value])  # density=False would make counts
    plt.title(f'{study_object} impacts per {declared_unit}')
    plt.ylabel('Probability (%)')
    plt.xlabel(f'GWP (kgCO2e/{declared_unit})')

    base_path = os.path.splitext(results_path)[0]
    plt.savefig(f'{base_path}.png')


def get_results(model, lcia_methods, counter=0, parameter_redefs=None):
    """
    Runs a simulation and returns a set of results to be stored.
    :param model_ref: reference to product system.
    :param lcia_methods: list of lcia methods to run calculations for.
    :param counter: counter from outer scope
    :param parameter_redefs: list of parameter redefinitions - this needs to be in OLCA format.
    :return: list of results
    """

    try:
        model_ref = client.get_descriptor(olca.ProductSystem, model.id)
    except IndexError:
        print("IndexError in OLCA calculation, probably because a product system was not set up correctly.")
        sys.exit()

    # reset all results
    gwp = np.nan
    gwp_be = np.nan
    gwp_bu = np.nan
    ap = np.nan
    ep = np.nan
    odp = np.nan
    pocp = np.nan
    gwp_ar5 = np.nan
    gwp_ef2 = np.nan
    gwp_cml = np.nan

    # Run simulations using each LCIA method and assign results to mapped variables
    if parameter_redefs is None:
        parameter_redefs = []
    print(f"Firing up OpenLCA simulator.")
    for lcia in lcia_methods:

        # print(f"\nSetting up OpenLCA simulator for {lcia}.")
        setup = olca.CalculationSetup(
            target=model_ref,
            impact_method=fetch_lcia_method(lcia),
            parameters=parameter_redefs,
            allocation=olca.AllocationType.USE_DEFAULT_ALLOCATION
        )

        result = client.calculate(setup)
        result.wait_until_ready()
        results = result.get_total_impacts()

        # print system setup for QA
        # print(f"Product system: {setup.target.name}\n"
        #       f"Impact method: {setup.impact_method.name}\n"
        #       f"Allocation method: {setup.allocation.name}\n"
        #       f"Amount: {ref_amount} {ref_unit}"
        #       )

        for r in results:
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Global warming']:
                gwp = r.amount
                gwp_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Global warming - biogenic emissions']:
                gwp_be = r.amount
                gwp_be_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Global warming - biogenic uptake']:
                gwp_bu = r.amount
                gwp_bu_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Acidification']:
                ap = r.amount
                ap_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Eutrophication']:
                ep = r.amount
                ep_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Ozone depletion']:
                odp = r.amount
                odp_unit = r.impact_category.ref_unit
            if lcia in ['TRACI 2.1', 'TRACI 2.1 (openIMPACT)'] \
                    and r.impact_category.name in ['Smog formation']:
                pocp = r.amount
                pocp_unit = r.impact_category.ref_unit
            if lcia in ['IPCC 2013 GWP 100a'] \
                    and r.impact_category.name in ['IPCC GWP 100a']:
                gwp_ar5 = r.amount
                gwp_ar5_unit = r.impact_category.ref_unit
            if lcia in ['EF Method (adapted)'] \
                    and r.impact_category.name in ['Climate change - fossil']:
                gwp_ef2 = r.amount
                gwp_ef2_unit = r.impact_category.ref_unit
            if lcia in ['CML-IA baseline'] \
                    and r.impact_category.name in ['Global warming (GWP100a)']:
                gwp_cml = r.amount
                gwp_cml_unit = r.impact_category.ref_unit
            else:
                pass
            # print(f"{r.impact_category.name} :: {r.amount} {r.impact_category.ref_unit}")

        # Dispose of simulator results before starting the next calculation setup and simulation.
        print(f"Completed analysis for: {results[0].impact_category.category}")
        result.dispose()
    result.dispose()

    impact_results = [round(gwp, 4), gwp_be, gwp_bu, ap, ep, odp, pocp,
                      round(gwp_ar5, 4),
                      round(gwp_ef2, 4),
                      round(gwp_cml, 4)]

    print(f'\nResult saved to csv. | Run {counter} gwp: {gwp:.2f} {gwp_unit} ({lcia_methods[0]})')

    return impact_results


"""CACHING FUNCTIONS"""

cache_process = dict()
cache_flows = dict()
cache_lcia = dict()
cache_ref_flows = dict()


def fetch_process_json(olca_uuid):
    """
    Caches process json to speed up loading of repeated processes.

    :param olca_uuid:
    :return: olca_json
    """
    # print(f"Checking for {olca_uuid} in cache.")
    if olca_uuid not in cache_process:
        cache_process[olca_uuid] = client.get(olca.Process, olca_uuid)
    #     print(f"Added {olca_uuid} to cache.")
    # else:
    #     print(f"{olca_uuid} already in cache.")

    return cache_process[olca_uuid]


def fetch_flow(olca_name):
    """
    Caches flow json to speed up loading of repeated processes.

    :param olca_name:
    :return: olca_json
    """
    # print(f"Checking for {olca_uuid} in cache.")
    if olca_name not in cache_flows:
        cache_flows[olca_name] = client.find(olca.Flow, olca_name)

    return cache_flows[olca_name]


def fetch_ref_flows(process):
    """
    Caches reference flow json to speed up loading of repeated processes.

    :param process:
    :return: olca_json
    """
    if process not in cache_ref_flows:
        for i in process.exchanges:
            if i.is_quantitative_reference:
                fname = i.flow.name
                fuuid = i.flow.id
                ftype = i.flow_property.name
                funit = i.unit.name
                fref = fetch_flow(fname)
        cache_ref_flows[process] = [fname, fuuid, ftype, funit, fref]

    return cache_ref_flows[process]


def fetch_lcia_method(olca_name):
    """
    Caches lcia method reference to speed up loading of methods.
    :param olca_name:
    :return: olca_ref
    """
    # print(f"Checking for {olca_uuid} in cache.")
    if olca_name not in cache_lcia:
        cache_lcia[olca_name] = client.find(olca.ImpactMethod, olca_name)

    return cache_lcia[olca_name]


if __name__ == "__main__": main()

