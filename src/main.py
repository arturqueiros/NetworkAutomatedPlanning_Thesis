import csv
import itertools
import random
from typing import List

import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import time
import io
import tempfile

import matplotlib.pyplot as plt
from collections import Counter
import math
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt


class KilometricPoint:
    """ This symbolizes a single Kpoint and its corresponding coverages for each station """

    def __init__(self, point_val: float, coverages: list[str]):
        self.point_val = point_val
        self.coverages = coverages


class DataPoint:
    """ [This is redudant, for the sake of testing]
        Each plot point: KilometricPoint and its maximum coverage for set of active interest points. """

    def __init__(self, pk: KilometricPoint, max_coverage: float):
        self.pk = pk
        self.max_coverage = max_coverage


class LineCoverage:
    """ This symbolizes the coverage value along the railway line"""

    def __init__(self, points: list[str], pk_values: list[float]):
        self.max_coverages = [datapoint.max_coverage for datapoint in points]
        self.count = 0
        self.points_below_ref = self.compute_points_below_ref(pk_values)

    def compute_points_below_ref(self, pk_values: list[float]):
        points_below_ref = 0
        for i in range(len(self.max_coverages)):
            if low_signal_ref >= self.max_coverages[i] >= lim_min_coverage:
                points_below_ref += 1
            if self.max_coverages[i] < lim_min_coverage or math.isnan(self.max_coverages[i]):
                self.count += 1
        return points_below_ref

    def get_percent_below_min(self):
        return (self.count / len(self.max_coverages)) * 100

    def get_percent_above_min(self):
        return 100 - ((self.count / len(self.max_coverages)) * 100)

    def get_percent_below_ref(self):
        return (self.points_below_ref / len(self.max_coverages)) * 100

    @property
    def get_max_coverages(self):
        return self.max_coverages


class Individual:
    """This symbolizes one individual - a random test case with random active/non-active stops."""

    def __init__(self, stopTypes: list[str]):
        self.stopTypes = stopTypes  # List of types of stops (Strings)
        # Using _ because nothing depends on that variable on the for
        self.individualStops = [random.randint(0, 1) for _ in stopTypes]
        self.id = 0
        # self.verifyAnchors()
        self.value_cost_function = 0
        self.individual_weight = 0
        self.individual_weight_percent = 0
        self.percent_below_min = 0
        self.percent_below_ref = 0
        self.max_extension_low_ref = 0
        self.num_station = 0
        self.num_halt = 0
        self.num_anchor = 0
        self.num_signal = 0
        self.num_level_crossing = 0
        self.num_other = 0
        self.weight_active_stations = 35
        self.weight_percent_below_min = 55
        self.weight_percent_low_ref = 20
        self.weight_extension = 25
        self.weight_station = 2
        self.weight_halt = 3
        self.weight_anchor = 1
        self.weight_level_crossing = 5
        self.weight_signal = 7
        self.weight_other = 10
        self.individual_coverage = []
        self.individual_pks = []

    def set_individualStops(self, individualStops: list[int]):
        self.individualStops = individualStops

    def get_Number_Active_Stops(self) -> int:
        return self.individualStops.count(1)  # Count the number of active stations

    # def set_individual_stops(self):
    #    worst_individual_stops = self.individualStops = [0] * len(self.stopTypes)
    #    return worst_individual_stops

    def get_Number_Stops(self) -> int:
        return len(self.stopTypes)

    def compute_cost_function(self, stations_active_number: int, below_min: float, low_ref: float, extension: float,
                              list_stations: list[int],
                              stopTypes: list[str]):
        halt = 0
        station = 0
        anchor = 0
        level_crossing = 0
        signal = 0
        other = 0
        # weight_sum = self.weight_active_stations + self.weight_percent_below_min + self.weight_percent_low_ref + self.weight_extension
        total_stops = self.get_Number_Stops()

        # num_activeStations = self.get_Number_Active_Stops() # TODO: refactor this method: stations active number can be accessed from within the class

        for index, stop in enumerate(self.individualStops):
            if stop:  # Check if stop is active
                term = stopTypes[index]
                match term:
                    case 'Station':
                        station += 1
                    case 'Halt':
                        halt += 1
                    case 'Level Crossing':
                        level_crossing += 1
                    case 'Anchor':
                        anchor += 1
                    case 'Signal':
                        signal += 1
                    case 'Other':
                        other += 1
        if anchor == 0 and stopTypes.count('Anchor') != 0:
            print('Anchor station must be always active in every solution!')
            exit()
        self.num_station = station
        self.num_halt = halt
        self.num_signal = signal
        self.num_anchor = anchor
        self.num_level_crossing = level_crossing
        self.num_other = other
        best_value = ((2 / total_stops) * self.weight_station) + ((1 / total_stops) * self.weight_halt) + \
                     ((1 / total_stops) * self.weight_anchor) + ((0 / total_stops) * self.weight_level_crossing) + \
                     ((0 / total_stops) * self.weight_signal) + ((0 / total_stops) * self.weight_other) + \
                     (0 * self.weight_percent_below_min) + (0 * self.weight_percent_low_ref) + (
                             0 * self.weight_extension)

        self.value_cost_function = ((self.num_station / total_stops) ** self.weight_station) + (
                (self.num_halt / total_stops) ** self.weight_halt) + \
                                   ((self.num_anchor / total_stops) ** self.weight_anchor) + (
                                           (self.num_level_crossing / total_stops) ** self.weight_level_crossing) + \
                                   ((self.num_signal / total_stops) ** self.weight_signal) + (
                                           (self.num_other / total_stops) ** self.weight_other) + \
                                   (below_min * self.weight_percent_below_min) + (
                                           low_ref * self.weight_percent_low_ref) + (
                                           extension * self.weight_extension)

        return self.value_cost_function

    @property
    def get_individualStops(self):
        return self.individualStops

    def __str__(self):
        return f"{self.individualStops}"


def read_coverages_file():
    stations = []
    pk_list = []
    pk_values = []
    with open('Coverage_Cascais.csv', 'r') as source_file:
        data_reader = csv.reader(source_file)

        for i, line in enumerate(data_reader):
            if i == 0:
                stations = [name for name in line]
                stations.pop(0)
            else:
                try:
                    pk_value = float(line[0])
                except ValueError:
                    print("Exception: Reached a pk that is not a float.")

                """ Append all coverage values. If the value is missing add a sufficiently low number instead. """
                pk_coverages = []
                for value in line:
                    try:
                        pk_coverages.append(float(value))
                    except ValueError:
                        if value == '':

                            pk_coverages.append(None)
                        else:
                            print(
                                "Exception: Reached a coverage that is not a float and is not an empty string either.")
                pk_coverages.pop(0)

                """ Create new KilometricPoint with its value and coverages for each interest point. """
                pk_list.append(KilometricPoint(pk_value, pk_coverages))
                pk_values.append(pk_value)
    return pk_values, stations, pk_list


def file_creation(filename1: str, filename2: str, generations_number: int):
    table_hearders_all_data = ["Individual", "Line Total Stops", "Cost Function Value", "Percentage Below Minimum", "Weight",
                               "Percentage Low Reference", "Weight", "Maximum Extension Low Reference", "Weight",
                               "No. Active Sites", "No. Station", "Weight", "No. Halt", "Weight", "No. Anchors",
                               "Weight", "No. Signal", "Weight", "No. Level_Crossing", "Weight", "No. Other",
                               "Weight", "Variable Cost Function Value"]
    table_hearders_best_individuals = ["Generation", "Individual", "Cost Function Value", "Percentage Below Minimum",
                                       "Percentage Low Reference",
                                       "Maximum Extension Low Reference", "No. Active Sites", "No. Station", "No. Halt",
                                       "No. Anchors",
                                       "No. Signal", "No. Level_Crossing", "No. Other", "Graphic"]

    # Creating 2 new workbook
    workbook1 = openpyxl.Workbook()
    workbook2 = openpyxl.Workbook()

    # Creating the number of pages according to the number of generations
    for i in range(generations_number):
        sheet_name = "Generation n." + str(i + 1)
        worksheet1 = workbook1.create_sheet(sheet_name)

        for j, header in enumerate(table_hearders_all_data):
            worksheet1.cell(row=1, column=j + 1, value=header)

    sheet_name1 = "Cost_Function Evolution"
    worksheet2 = workbook2.create_sheet(sheet_name1)
    sheet_name2 = "Cost_Function Generation"
    worksheet2 = workbook2.create_sheet(sheet_name2)

    for i, header in enumerate(table_hearders_best_individuals):
        worksheet2.cell(row=1, column=i + 1, value=header)

    workbook1.remove(workbook1[workbook1.sheetnames[0]])
    workbook1.save(filename1)
    workbook2.remove(workbook2[workbook2.sheetnames[0]])
    workbook2.save(filename2)

    return workbook1, workbook2


# def read_cov_file():
#     stations = []
#     pk_list = []
#     pk_values = []
#     with open('Coverage_Cascais.csv', 'r') as source_file:
#         data_reader = csv.reader(source_file)
#
#         for i, line in enumerate(data_reader):
#             if i == 0:
#                 stations = [name for name in line]
#                 stations.pop(0)
#             else:
#                 try:
#                     pk_value = float(line[0])
#                 except ValueError:
#                     print("Exception: Reached a pk that is not a float.")
#
#                 """ Append all coverage values. If the value is missing add a sufficiently low number instead. """
#                 pk_coverages = []
#                 for value in line:
#                     try:
#                         pk_coverages.append(float(value))
#                     except ValueError:
#                         if value == '':
#
#                             pk_coverages.append(None)
#                         else:
#                             print(
#                                 "Exception: Reached a coverage that is not a float and is not an empty string either.")
#                 pk_coverages.pop(0)
#
#                 """ Create new KilometricPoint with its value and coverages for each interest point. """
#                 pk_list.append(KilometricPoint(pk_value, pk_coverages))
#                 pk_values.append(pk_value)
#
#     stations_active_values = read_input_active_stations(type_read())
#     interesting_pks = identify_station_pks(stations, stations_active_values[0])
#
#     datapoints = processDataPoints(stations_active_values[0], pk_list)
#
#     processStationPoints(interesting_pks, datapoints)
#
#     return datapoints, stations_active_values, pk_values


"""
def compute_values(points: list[str]):
    total_points = len(max_coverages)

    count = 0
    nancount = 0
    points_below_ref = 0
    points_in_ref: list[KilometricPoint] = []

    for i in range(len(max_coverages)):
        if max_coverages[i] < lim_min_coverage or math.isnan(max_coverages[i]):
            count += 1

    dist_points = []
    total_distance = 0
    for i in range(len(max_coverages)):
        if low_signal_ref >= max_coverages[i] >= lim_min_coverage:
            points_below_ref += 1
            points_in_ref.append(pk_values[i])
            dist_points.append(pk_values[i])

        else:
            total_distance += dist_calc(dist_points)
            dist_points.clear()

    percent_below_min = (count / total_points) * 100
    percent_above_min = 100 - ((count / total_points) * 100)
    percent_below_ref = (points_below_ref / total_points) * 100

    return max_coverages, percent_below_min, percent_above_min, percent_below_ref
"""


# def read_input_active_stations(types: list[str]):
#     list = initIndividual(types)
#     number_active_stations = list[0].count('1')
#     return list, number_active_stations


def identify_station_pks(stations_names: list[str], active_stations: list[str]):
    interesting_pks = []
    with open('Cascais Elements_pks.csv') as source_file:
        data_reader2 = csv.reader(source_file)
        station_names_and_pks = [station for station in data_reader2]
        station_names_and_pks.pop(0)

        for indx, station_status in enumerate(active_stations):
            if (station_status):
                station_name_1 = stations_names[indx]
                for station_name_and_pk in station_names_and_pks:
                    if station_name_1 == station_name_and_pk[0]:
                        interesting_pks.append(station_name_and_pk[1])
    return interesting_pks


def processDataPoints(station_list: list[str], pks: list[KilometricPoint]):
    data_points = []
    for pk in pks:

        max_coverage = -1200  # If a pk has all values missing
        # np.max(pk.coverages[pk])
        for j, station in enumerate(station_list):  # Checks the highest coverage value
            if pk.coverages[j] is not None:
                if station and pk.coverages[j] > max_coverage:  # Station is active
                    max_coverage = pk.coverages[j]
        if max_coverage != -1200:
            data_points.append(DataPoint(pk, max_coverage))
        else:
            data_points.append(DataPoint(pk, float('nan')))
    return data_points


def processStationPoints(active_station_pks: list[str], datapoints: list[DataPoint]):
    for station_pk in active_station_pks:
        for indx, datapoint in enumerate(datapoints):
            if datapoint.pk.point_val == float(station_pk):
                adjacent_value = None
                offset = 1
                while (
                        adjacent_value is None):  # find the closest adjancent value, that actually has a value!
                    # Perhaps change this to find the max adjacent value instead...
                    direction = offset % 2
                    if direction == 1:
                        if indx + offset < len(datapoints):
                            try:
                                adjacent_value = datapoints[indx + offset].max_coverage  # check values on right
                            except:
                                print("index", indx)
                                print("offset", offset)
                                print("len(datapoints)", len(datapoints))
                        elif indx - offset >= 0:
                            adjacent_value = datapoints[indx - offset].max_coverage  # check values on right
                        else:
                            raise Exception
                    else:
                        if indx - offset >= 0:
                            adjacent_value = datapoints[indx - offset].max_coverage  # check values on left
                        elif indx + offset <= len(datapoints):
                            adjacent_value = datapoints[indx + offset].max_coverage  # check values on right
                        else:
                            raise Exception
                    offset += 1
                datapoint.max_coverage = adjacent_value
                break
            elif datapoint.pk.point_val > float(station_pk):
                break
            # if(datapoint.pk > station_pk): # if we passed the point value, we won't be finding it ahead either
            #  break


"""
def dist_calc(seq_points_in_ref: list[int]):
    receives array with pk and calculates distance between first and last point
    if len(seq_points_in_ref) == 0 or len(seq_points_in_ref) == 1:
        return 0

    return seq_points_in_ref[len(seq_points_in_ref) - 1] - seq_points_in_ref[0]
"""


def compute_max_extension(coverage_points: list[float], pk_reference: list[float]):
    point_first = None
    point_last = None
    max_extension = 0
    for index, coverage_point in enumerate(coverage_points):
        if lim_min_coverage <= coverage_point < low_signal_ref:  # inside limits
            if point_first is None:
                point_first = pk_reference[index]
            if point_first is not None and index + 1 >= len(coverage_points):  # in case its the last point
                point_last = pk_reference[index]
                max_temp = compute_extension_distance(point_first, point_last)
                max_extension = max_temp if max_temp > max_extension else max_extension
        if (lim_min_coverage > coverage_point or coverage_point >= low_signal_ref) and (
                point_first is not None):  # outside limits
            point_last = pk_reference[index - 1]
            max_temp = compute_extension_distance(point_first, point_last)
            max_extension = max_temp if max_temp > max_extension else max_extension
            point_first = None
            point_last = None

    return max_extension


def compute_extension_distance(point_first: float, point_last: float):
    max_distance_temp = point_last - point_first
    max_distance = 0
    if max_distance_temp > max_distance:
        max_distance = max_distance_temp
    return max_distance


"""
def cost_func(stations_active_number: int, below_min: float, low_ref: float, extension: float, list_stations: list[int],
              priorities: list[str]):
    height_stations = 100
    height_percent_below_min = 150
    height_percent_low_ref = 40
    height_extension = 30
    heights_priorities = 15
    height_station = 1
    height_halt = 2
    height_level_crossing = 4
    height_signal = 7
    height_other = 10
    halt = 0
    station = 0
    anchor = 0
    level_crossing = 0
    signal = 0
    other = 0
    height_sum = height_stations + height_percent_below_min + height_percent_low_ref + height_extension + heights_priorities
    total_stations = len(list_stations)

    for index, this_station in enumerate(list_stations):
        if this_station == '1':
            if priorities[index] == 'Station':
                station += 1
            elif priorities[index] == 'Halt':
                halt += 1
            elif priorities[index] == 'Anchor':
                anchor += 1
            elif priorities[index] == 'Level Crossing':
                level_crossing += 1
            elif priorities[index] == 'Signal':
                signal += 1
            elif priorities[index] == 'Other':
                other += 1
    if anchor == 0 and priorities.count('Anchor') != 0:
        print('Anchor station must be always active in every solution!')
        exit()

    cost_func_value = (
                              stations_active_number / total_stations) * height_stations / height_sum + below_min * \
                      height_percent_below_min / height_sum + (
                              low_ref * height_percent_low_ref) / height_sum + (
                              extension * height_extension) / height_sum + (
                              (station ^ height_station) + (halt ^ height_halt) + (
                              level_crossing ^ height_level_crossing) + (signal ^ height_signal) + (
                                      other ^ height_other)) * heights_priorities / height_sum
    return cost_func_value
"""


def read_stopTypes():
    types = []
    with open('Cascais Elements.csv') as types_file:
        data_reader = csv.reader(types_file)
        next(data_reader)
        for row in data_reader:
            types.append(row[1])
    return types


# def initIndividual(types: list[str]):
#     size = len(types)
#     individual = []
#     # Generate a random string of 0s and 1s wich will represent the individual
#     for i in range(size):
#         individual.append(str(random.randint(0, 1)))
#
#     print("Array:", individual)
#     # Verify if there is any Anchor on the stations
#     for index, type in enumerate(types):
#         if type == 'Anchor' and not individual[index]:
#             # Force the position with the Anchor to be active (1)
#             individual[index] = '1'
#
#     value_cost_function = cost_func(stations_active_values[1], percent_below_min, percent_below_ref,
#                                     points_extension(max_coverages, pk_values), stations_active_values[0], type_read())
#
#     print('Valor funcão custo:', value_cost_function)
#
#     return individual, value_cost_function

def create_population(size_population: int, stop_types: list[str]):
    # Generates the number os individuals according to the size of the population
    active_station_dict = {}
    population = []
    i = 0
    while i < size_population:
        individual = Individual(stop_types)
        individualStops = individual.get_individualStops
        if str(individualStops) not in active_station_dict:
            active_station_dict[str(individualStops)] = i
            population.append(individual)
            i += 1
        individual.id = i

    return population


def verify_anchors(population: list[Individual]):
    for individual in population:
        for index, stop in enumerate(individual.individualStops):
            # checks if the archor station is not set to 1
            if individual.stopTypes[index] == 'Anchor' and stop == 0:
                # Force the position with the Anchor to be active (1)
                individual.individualStops[index] = 1
    return population


def compute_population_data(population: list[Individual]):
    # print("População:")
    population = verify_anchors(population)
    for individual in population:
        num_active_stops = individual.get_Number_Active_Stops()
        individualStops = individual.get_individualStops
        activeStops_pks = identify_station_pks(stations, individualStops)
        datapoints = processDataPoints(individualStops, pk_list)
        processStationPoints(activeStops_pks, datapoints)
        line_coverage = LineCoverage(datapoints, pk_values)

        percent_below_min = line_coverage.get_percent_below_min()
        percent_below_ref = line_coverage.get_percent_below_ref()
        # percent_above_min = line_coverage.get_percent_above_min()
        max_coverages = line_coverage.get_max_coverages

        # Compute the consecutive max entension between the pk values below the minimum limit
        max_extension = compute_max_extension(max_coverages, pk_values)

        # Compute the cost function value for the individual
        individual_cost_function_value = individual.compute_cost_function(num_active_stops, percent_below_min,
                                                                          percent_below_ref, max_extension, stations,
                                                                          stop_Types)
        individual.percent_below_min = percent_below_min
        individual.percent_below_ref = percent_below_ref
        individual.max_extension_low_ref = max_extension
        individual.individual_coverage = line_coverage
        individual.individual_pks = pk_values
        # print("\t", individualStops, num_active_stops, individual_cost_function_value)
    return population


def rank_individuals(population: list[Individual]):
    return population.sort(key=lambda individual: individual.value_cost_function, reverse=False)


def selection(selection_population: list[Individual], eliteSize: int):
    population_total_cost_val_sum = 0
    cumulative_weights = [0]
    parents = []
    j = 0
    elite_index = 0
    elite = []
    for i in range(eliteSize):
        selection_population.remove(selection_population[i])
    for individual in selection_population:
        population_total_cost_val_sum += individual.value_cost_function
        normalized_weights = individual.value_cost_function
        cumulative_weights.append(cumulative_weights[j] + normalized_weights)
        individual.individual_weight_percent = (cumulative_weights[j + 1] / population_total_cost_val_sum) * 100
        j += 1
    cumulative_weights.pop(0)
    # Roulette selection method, the number of parents obtain from this selection must never be smaller than half of the population size
    max_value = max(cumulative_weights)
    for count in range(len(selection_population)):
        random_number = random.uniform(0, max_value)
        for score in cumulative_weights:
            if random_number <= score:
                parents.append(selection_population[count])
                cumulative_weights.remove(score)
                break
    # parents = list(set(parents))
    # print("Selection:", parents)

    return parents


def crossover(parents: list[Individual], crossover_probability: float):
    crossover_population = []
    random.shuffle(parents)
    parents_check = []
    parents_check.extend(parents)
    # Count the number of pairs for crossover
    number_pairs = math.floor((len(parents) / 2))

    # Get the size of a chromossome
    chromossome_size = len(parents[0].individualStops)

    for pair in range(number_pairs):
        if random.random() < crossover_probability:
            length = len(parents)
            first_parent_index = random.randrange(0, length)
            second_parent_index = random.randrange(0, length)
            while first_parent_index == second_parent_index:
                second_parent_index = random.randrange(0, length)

            # start = random.randrange(chromossome_size)
            stop = round(chromossome_size / 2)
            # if start > stop:
            #    start, stop = stop, start

            # Create child chromosomes by combining sections of the parents' chromosomes
            first_parent = parents[first_parent_index]
            second_parent = parents[second_parent_index]
            first_child = Individual(list(first_parent.individualStops))
            first_child.set_individualStops(first_parent.individualStops[0:stop] + second_parent.individualStops[stop:])
            second_child = Individual(list(second_parent.individualStops))
            second_child.set_individualStops(
                second_parent.individualStops[0:stop] + first_parent.individualStops[stop:])

            # Remove the parents from the pool
            parents.remove(first_parent)
            parents.remove(second_parent)

            crossover_population.append(first_child)
            crossover_population.append(second_child)
            crossover_population = compute_population_data(crossover_population)

    """print("População Pré mutação:")
    for crossover_individual in crossover_population:
        print("\t", crossover_individual)
    """

    return crossover_population


def mutation(mutation_population: list[Individual], mutation_probability: float):
    # This method will apply a bit swap to one of the individuals on the crossover population
    # print("População pós mutação:")
    population = []
    population.extend(mutation_population)
    for individual in mutation_population:
        rand_numb = round(random.uniform(0, 1), 2)
        if rand_numb <= mutation_probability:
            index_1 = random.randrange(len(individual.individualStops))
            # index_2 = random.randrange(len(individual.individualStops))

            if individual.stopTypes[index_1] == 'Anchor':
                individual.individualStops[index_1] = 1
            else:
                if individual.individualStops[index_1] == 1:
                    individual.individualStops[index_1] = 0
                else:
                    individual.individualStops[index_1] = 1

            """temp = individual.individualStops[index_1]
            individual.individualStops[index_1] = individual.individualStops[index_2]
            individual.individualStops[index_2] = temp"""

        # print("\t", individual)
    mutation_population = compute_population_data(mutation_population)
    # population = population + mutation_population
    return mutation_population


def generations_creation(population_size: int, stop_Types: list[str], number_generations: int, eliteSize: int,
                         crossover_probability: float, mutation_probability: float, data_file: Workbook):
    best_individuals = []
    elite_individuals = []
    population = create_population(population_size, stop_Types)
    # population[0].individualStops = [0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0]
    computed_population = compute_population_data(population)
    rank_individuals(computed_population)
    all_populations_data = []
    index = 0

    for generation in range(number_generations):
        number_elite = len(elite_individuals)
        if number_elite == 0:
            elite = get_elite(computed_population, eliteSize)
            parents = selection(population, eliteSize)

        new_population_size = population_size - number_elite
        if number_elite > 0:
            elite = get_elite(final_population, eliteSize)
            parents = selection(final_population, eliteSize)

        crossover_population = crossover(parents, crossover_probability)
        after_crossover_population = crossover_population + elite
        # Checks the new_population size and if its lower than initial population size, fills the population with new random individuals
        filled_population = fill_population(after_crossover_population, population_size, stop_Types)
        # Generate a random number between 0 and 1 to see if the mutation will occur
        final_population = mutation(filled_population, mutation_probability)
        rank_individuals(final_population)
        elite_individuals = get_elite(final_population, eliteSize)
        best_individuals.append(get_best_individual(final_population[0]))
        all_data(final_population, generation, data_file)
        generation_data = []
        for index in range(len(final_population)):
            generation_data.append(final_population[index])
        all_populations_data.append(generation_data)
        parents = []
        after_crossover_population = []

    return best_individuals, all_populations_data


def fill_population(new_population: list[Individual], population_size: int, stop_Types: list[str]):
    this_population_size = population_size - len(new_population)
    filling_population = create_population(this_population_size, stop_Types)
    filling_population = compute_population_data(filling_population)
    new_population = new_population + filling_population

    return new_population


def get_elite(new_population: list[Individual], eliteSize: int):
    # Getting the elite individuals to go into the next gen
    elite_individuals = []
    number_elite = 0
    for individual in new_population:
        if number_elite < eliteSize:
            elite_individuals.append(individual)
            number_elite += 1
        else:
            break
    return elite_individuals


def all_data(new_population: list[Individual], generation: int, all_data_file: Workbook):
    sheet_name1 = "Generation n." + str(generation + 1)
    worksheet = all_data_file[sheet_name1]

    for k, individual in enumerate(new_population):
        row = k + 2  # Offset by 1 to account for headers
        worksheet.cell(row=row, column=1, value=str(individual.individualStops))
        worksheet.cell(row=row, column=2, value=float(individual.get_Number_Stops()))
        worksheet.cell(row=row, column=3, value=float(individual.value_cost_function))
        worksheet.cell(row=row, column=4, value=float(individual.percent_below_min))
        worksheet.cell(row=row, column=5, value=float(individual.weight_percent_below_min))
        worksheet.cell(row=row, column=6, value=float(individual.percent_below_ref))
        worksheet.cell(row=row, column=7, value=float(individual.weight_percent_low_ref))
        worksheet.cell(row=row, column=8, value=float(individual.max_extension_low_ref))
        worksheet.cell(row=row, column=9, value=float(individual.weight_extension))
        worksheet.cell(row=row, column=10, value=individual.get_Number_Active_Stops())
        worksheet.cell(row=row, column=11, value=individual.num_station)
        worksheet.cell(row=row, column=12, value=individual.weight_station)
        worksheet.cell(row=row, column=13, value=individual.num_halt)
        worksheet.cell(row=row, column=14, value=individual.weight_halt)
        worksheet.cell(row=row, column=15, value=individual.num_anchor)
        worksheet.cell(row=row, column=16, value=individual.weight_anchor)
        worksheet.cell(row=row, column=17, value=individual.num_signal)
        worksheet.cell(row=row, column=18, value=individual.weight_signal)
        worksheet.cell(row=row, column=19, value=individual.num_level_crossing)
        worksheet.cell(row=row, column=20, value=individual.weight_level_crossing)
        worksheet.cell(row=row, column=21, value=individual.num_other)
        worksheet.cell(row=row, column=22, value=individual.weight_other)
        cost_function_formula = f'(((K{row} / B{row}) ^ L{row}) + ((M{row} / B{row}) ^ N{row}) + ((O{row} / B{row}) ^ P{row}) + ((Q{row} / B{row}) ^ R{row}) + ((S{row} / B{row}) ^ T{row}) + ((U{row} / B{row}) ^ V{row}) + (D{row} * E{row}) + (F{row} * G{row}) + (H{row} * I{row}))'
        """
            f'(({worksheet.cell(row=row, column=10).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=12).coordinate}) + '  # num_station / total_stops
            f'(({worksheet.cell(row=row, column=13).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=12).coordinate}) + '  # num_halt / total_stops
            f'(({worksheet.cell(row=row, column=15).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=16).coordinate}) + '  # num_anchor / total_stops
            f'(({worksheet.cell(row=row, column=19).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=20).coordinate}) + '  # num_level_crossing / total_stops
            f'(({worksheet.cell(row=row, column=17).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=18).coordinate}) + '  # num_signal / total_stops
            f'(({worksheet.cell(row=row, column=21).coordinate} / {worksheet.cell(row=row, column=2).coordinate}) ** '
            f'{worksheet.cell(row=row, column=22).coordinate}) + '  # num_other / total_stops
            f'({worksheet.cell(row=row, column=4).coordinate} * {worksheet.cell(row=row, column=5).coordinate}) + '  # below_min * weight_percent_below_min
            f'({worksheet.cell(row=row, column=6).coordinate} * {worksheet.cell(row=row, column=7).coordinate}) + '  # low_ref * weight_percent_low_ref
            f'({worksheet.cell(row=row, column=8).coordinate} * {worksheet.cell(row=row, column=9).coordinate})'
        # extension * weight_extension
        )"""
        formula = "=" + cost_function_formula
        worksheet.cell(row=row, column=24).value = formula
        # worksheet.cell(row=row, column=22, value=cost_function_formula.format(row=row))
    workbook1.save(filename1)
    all_data_file.save(filename1)


def results_data(best_individuals: list[Individual], best_individuals_file: Workbook, generation: list[int],
                 best_individual_cost_value: list[int], population_size: int, number_generations: int,
                 crossover_prob: float, mutatio_prob: float):
    sheet_name1 = "Cost_Function Evolution"
    worksheet1 = best_individuals_file[sheet_name1]
    fig1, ax = plt.subplots()
    generation_evolution_plot(generation, best_individual_cost_value)

    # Convert plot to image
    image_stream = io.BytesIO()
    plt.savefig(image_stream, format='png')
    plt.close(fig1)

    # Insert the image into the Excel worksheet
    image_stream.seek(0)
    img = Image(image_stream)
    img.width = 300  # Adjust the width of the image if needed
    img.height = 220  # Adjust the height of the image if needed
    worksheet1.add_image(img, f'C{4}')
    worksheet1.cell(row=5, column=10, value=f"No. Individuals: {population_size}")
    worksheet1.cell(row=6, column=10, value=f"No. Generations: {number_generations}")
    worksheet1.cell(row=7, column=10, value=f"Crossover Prob.: {crossover_prob}")
    worksheet1.cell(row=8, column=10, value=f"Mutation Prob.: {mutatio_prob}")

    sheet_name2 = "Cost_Function Generation"
    worksheet2 = best_individuals_file[sheet_name2]
    for k, individual in enumerate(best_individuals):
        row = k + 2
        worksheet2.cell(row=row, column=1, value=k + 1)
        worksheet2.cell(row=row, column=2, value=str(individual.individualStops))
        worksheet2.cell(row=row, column=3, value=float(individual.value_cost_function))
        worksheet2.cell(row=row, column=4, value=float(individual.percent_below_min))
        worksheet2.cell(row=row, column=5, value=float(individual.percent_below_ref))
        worksheet2.cell(row=row, column=6, value=float(individual.max_extension_low_ref))
        worksheet2.cell(row=row, column=7, value=individual.get_Number_Active_Stops())
        worksheet2.cell(row=row, column=8, value=individual.num_station)
        worksheet2.cell(row=row, column=9, value=individual.num_halt)
        worksheet2.cell(row=row, column=10, value=individual.num_anchor)
        worksheet2.cell(row=row, column=11, value=individual.num_signal)
        worksheet2.cell(row=row, column=12, value=individual.num_level_crossing)
        worksheet2.cell(row=row, column=13, value=individual.num_other)

        # Generate the plot
        fig2, ax = plt.subplots()
        cov_data_plot(generation=k, best_individuals=best_individuals)

        # Convert plot to image
        image_stream = io.BytesIO()
        plt.savefig(image_stream, format='png')
        plt.close(fig2)

        # Insert the image into the Excel worksheet (Column 14 - Column N)
        image_stream.seek(0)
        img = Image(image_stream)
        worksheet2.column_dimensions['N'].width = 40
        worksheet2.row_dimensions[row].height = 170
        img.width = 300  # Adjust the width of the image if needed
        img.height = 220  # Adjust the height of the image if needed
        worksheet2.add_image(img, f'N{row}')

        # worksheet.cell(row=row, column=14, value=cov_data_plot(k, best_individuals))
    best_individuals_file.save(filename2)


def get_best_individual(best_individual: Individual):
    """all_best_individual = [(best_individual.individualStops, best_individual.value_cost_function, generation,
                            best_individual.individual_pks, best_individual.individual_coverage, best_individual.num_anchor, best_individual.num_station,
                            best_individual.num_halt, best_individual.num_signal, best_individual.num_level_crossing, best_individual.num_other,
                            best_individual.get_Number_Active_Stops())]"""
    all_best_individual = best_individual
    return all_best_individual


def cov_data_plot(generation: int, best_individuals: list[Individual]):
    # for generation in range(number_generations):
    pk_axis = best_individuals[generation].individual_pks
    coverage_axis_aux = best_individuals[generation].individual_coverage
    coverage_axis = coverage_axis_aux.get_max_coverages
    plt.figure(generation + 1)
    plt.plot(pk_axis, coverage_axis)
    plt.xlim(0, max(pk_values))
    plt.axhline(y=lim_min_coverage, color='r')
    plt.axhline(y=low_signal_ref, color='y')
    plt.xlabel("Distance (pk)")
    plt.ylabel("Coverage (dBm)")
    title = f"Coverage Map of the best individual of Generation number {generation + 1}"
    plt.title(title)


def generation_evolution_plot(generation: list[int], best_individual_cost_value: list[int]):
    plt.figure(1)
    plt.plot(generation, best_individual_cost_value)
    plt.xlabel("Generation")
    plt.ylabel("Best individual cost function value")
    plt.title("Cost Function Value evolution of the best individual")
    plt.xticks(range(min(generation), max(generation) + 1, 1))


if __name__ == "__main__":
    start_time = time.time()
    filename1 = "data_storage.xlsx"
    filename2 = "Result_Test.xlsx"
    lim_min_coverage = -95
    low_signal_ref = -85
    # This will be the initial size, the population after selection crossover and
    # mutation will have a random size based on the selection occurance
    population_size = 20
    number_generations = 10
    crossover_probability = 0.50
    mutation_probability = 0.01
    best_individuals = []
    all_cost_function_data = []

    # Create the xlsx file to store the data from all the generations
    workbook1, workbook2 = file_creation(filename1, filename2, number_generations)

    # Size of elite population should never be higher than 5% of the population size
    eliteSize = round(0.01 * population_size)
    if eliteSize == 0:
        eliteSize = 1

    # Read external files
    pk_values, stations, pk_list = read_coverages_file()
    stop_Types = read_stopTypes()

    """
    print("Ranked individuals:")
    for ranked_individuals in computed_population:
        print("\t Id:",ranked_individuals.id, ranked_individuals, ranked_individuals.value_cost_function)
    """
    best_individuals, all_population = generations_creation(population_size, stop_Types, number_generations,
                                                            eliteSize,
                                                            crossover_probability,
                                                            mutation_probability, workbook1)

    x = list(itertools.chain.from_iterable([[i + 1] * len(row) for i, row in enumerate(all_population)]))
    y = list(itertools.chain.from_iterable(all_cost_function_data))
    best_individual_cost_value = []
    generation = []
    gen = 1
    for individual in best_individuals:
        best_individual_cost_value.append(individual.value_cost_function)
        generation.append(gen)
        gen += 1

    end_time = time.time()
    elapsed_time = end_time - start_time
    print("Elapsed time: ", elapsed_time)

    results_data(best_individuals, workbook2, generation, best_individual_cost_value, population_size,
                 number_generations, crossover_probability, mutation_probability)

"""
    plt.figure(1)
    # plt.scatter(x, y, c=np.arange(len(y)), cmap='viridis')
    plt.plot(generation, best_individual_cost_value)
    plt.xlabel("Generation")
    plt.ylabel("Best individual cost function value")
    plt.title("Cost Function Value evolution of the best individual")
    plt.xticks(range(min(generation), max(generation) + 1, 1))
    # plt.yticks([i / 10 for i in range(int(min(best_individual_cost_value)), int(max(best_individual_cost_value) * 10) + 1, 1)])
    # plt.text(0.95, 0.05, elapsed_time, ha="right", va="bottom", transform=plt.gca().transAxes)
    plt.show()
    
    plt.figure(2)
    plt.scatter(x, y, c=np.arange(len(y)), cmap='viridis')
    plt.plot(generation, best_individual_cost_value)
    plt.xlabel("Generation")
    plt.ylabel("Best individual cost function value")
    plt.title("Cost Function Value evolution of the best individual")
    plt.xticks(range(min(generation), max(generation) + 1, 1))
    # plt.yticks([i / 10 for i in range(int(min(best_individual_cost_value)), int(max(best_individual_cost_value) * 10) + 1, 1)])
    # plt.text(0.95, 0.05, elapsed_time, ha="right", va="bottom", transform=plt.gca().transAxes)
    plt.show()
   
    # Create a single individual
    individual1 = Individual(stop_Types)
    activeStops = individual1.get_activeStops
    
    #Get number of active Stops
    num_active_stops = individual1.get_Number_Active_Stops()

    activeStops_pks = identify_station_pks(stations, activeStops)
    datapoints = processDataPoints(activeStops, pk_list)
    processStationPoints(activeStops_pks, datapoints)

    line_coverage = LineCoverage(datapoints, pk_values)

    percent_below_min = line_coverage.get_percent_below_min()
    percent_below_ref = line_coverage.get_percent_below_ref()
    percent_above_min = line_coverage.get_percent_above_min()
    max_coverages = line_coverage.get_max_coverages

    # Compute the consecutive max entension between the pk values below the minimum limit
    max_extension = compute_max_extension(max_coverages, pk_values)

    # Compute the cost function value for the individual
    individual_cost_function_value = individual1.compute_cost_function(num_active_stops, percent_below_min, percent_below_ref, max_extension, stations,
                                      stop_Types)

    print('Número de estações ativas:', num_active_stops)
    print('Percentagem de pontos acima do limite mínimo de cobertura:', percent_above_min, '%')
    print('Percentagem de pontos em low ref:', percent_below_ref, '%')
    print('Maior extensão de pontos em low ref:', max_extension)
    print('Valor da função de custo:', individual_cost_function_value)

    # sizeofPopulation = 100
    # initPopulation(sizeofPopulation,types)

    # Plot data in a single plot
    plt.figure(1)
    plt.plot(pk_values, max_coverages)
    plt.xlim(0, max(pk_values))
    plt.axhline(y=lim_min_coverage, color='r')
    plt.axhline(y=low_signal_ref, color='y')
    plt.xlabel("Distance (pk)")
    plt.ylabel("Coverage (dBm)")
    plt.show()
"""
