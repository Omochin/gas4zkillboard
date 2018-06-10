import os
import json
import yaml
import zipfile
from collections import OrderedDict


def unzip_sde():
    for filename in os.listdir():
        base, ext = os.path.splitext(filename)
        if ext == '.zip' and 'sde-' in base and '-TRANQUILITY' in base:
            with zipfile.ZipFile(filename) as zfile:
                zfile.extractall()
            break


def generate_type_ids_json():
    # 6: Ship
    # 7: Module
    # 8: Charge
    # 18: Drone
    # 22: Deployable
    # 23: Starbase
    # 46: Orbitals
    # 65: Structure
    # 66: Structure Module
    # 87: Fighter
    include_category_ids = [6, 7, 8, 18, 22, 23, 46, 65, 66, 87]
    include_group_ids = []
    type_dict = OrderedDict()

    base_path = os.path.join('.', 'sde', 'fsd')
    with open(os.path.join(base_path, 'groupIDs.yaml')) as file:
        for i, items in yaml.load(file).items():
            if items['categoryID'] in include_category_ids:
                include_group_ids.append(i)

    with open(os.path.join(base_path, 'typeIDs.yaml')) as file:
        for i, items in yaml.load(file).items():
            if items['groupID'] not in include_group_ids:
                continue

            type_id = str(i)
            try:
                type_name = str(items['name']['en'])
            except KeyError:
                type_name = ''

            type_dict[type_id] = type_name

    with open('type_ids.json', 'w', encoding='utf-8') as file:
        json.dump(type_dict, file)


def generate_universe_ids_json():
    def get_values(path, *names):
        values = {}
        with open(path) as file:
            target_yaml = yaml.load(file)
            for name in names:
                values[name] = target_yaml[name]
        return values

    region_dict = OrderedDict()
    constellation_dict = OrderedDict()
    solar_system_dict = OrderedDict()

    base_path = os.path.join('.', 'sde', 'fsd', 'universe')
    for universe_name in os.listdir(base_path):
        universe_path = os.path.join(base_path, universe_name)

        for region_name in os.listdir(universe_path):
            region_path = os.path.join(universe_path, region_name)

            if not os.path.isdir(region_path):
                continue

            region_values = get_values(
                os.path.join(region_path, 'region.staticdata'),
                'regionID'
            )
            region_id = str(region_values['regionID'])
            region_dict[region_id] = region_name

            for constellation_name in os.listdir(region_path):
                constellation_path = os.path.join(region_path, constellation_name)

                if not os.path.isdir(constellation_path):
                    continue

                constellation_values = get_values(
                    os.path.join(constellation_path, 'constellation.staticdata'),
                    'constellationID'             
                )
                constellation_id = str(constellation_values['constellationID'])
                constellation_dict[constellation_id] = OrderedDict(
                    name=constellation_name,
                    region_id=region_id,
                )

                for solar_system_name in os.listdir(constellation_path):
                    solar_system_path = os.path.join(constellation_path, solar_system_name)

                    if not os.path.isdir(solar_system_path):
                        continue

                    print(solar_system_path)

                    solar_system_values = get_values(
                        os.path.join(solar_system_path, 'solarsystem.staticdata'),
                        'solarSystemID',
                        'security'                 
                    )
                    solar_system_id = str(solar_system_values['solarSystemID'])
                    security = round(solar_system_values['security'], 1)
                    solar_system_dict[solar_system_id] = [
                        solar_system_name,
                        security,
                        region_id
                    ]

    with open('region_ids.json', 'w', encoding='utf-8') as file:
        json.dump(region_dict, file)

#    with open('constellation_ids.json', 'w', encoding='utf-8') as file:
#        json.dump(constellation_dict, file)

    with open('solar_system_ids.json', 'w', encoding='utf-8') as file:
        json.dump(solar_system_dict, file)


def main():
    unzip_sde()
    generate_type_ids_json()
    generate_universe_ids_json()


if __name__ == '__main__':
    main()
