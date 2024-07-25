import random
import pandas as pd
import copy

from enum import Enum
from dataclasses import dataclass

DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
TIME_SLOTS = ("10:10-11:00", "11:05-11:55", "12:00-12:45", "1:15-2:05", "2:10-3:00", "3:05-3:55", "4:00-4:45")
BREAK_FROM = 3
NO_THEORY_FROM = 6

ClassType = Enum("ClassType", ['NA', 'MAJOR', 'MINOR', 'MDS', 'ENVS'])


@dataclass
class Slot:
    room_id: str = None
    classtype: ClassType = ClassType.NA
    is_practical: bool = False
    cont: bool = False

    def __str__(self):
        if self.classtype == ClassType.NA:
            return "X"
        return self.classtype.name + "(" + ("P" if self.is_practical else "T") + ")" + ("*" if self.cont else "") + (f" {self.room_id}" if self.room_id != None else "")


@dataclass
class Room:
    room_id: str
    capacity: int
    hasAC: bool
    hasAV: bool
    floor: int


class Batch:
    def __init__(self, dept, prgm, sem, hasoffday, strength, ths, prs):
        self.dept = dept
        self.prgm = prgm
        self.sem = sem
        self.hasoffday = hasoffday
        self.strength = strength
        self.ths = ths
        self.prs = prs
        self.grid = [[Slot() for _ in TIME_SLOTS] for _ in DAYS]
        self.offday = None

    def __allocate(self, day, classtype, is_practical, cons, start, end, batches, rooms):
        i = start
        while i <= end:
            flag = True
            while i <= end and self.grid[day][i].classtype != ClassType.NA:
                i += 1
            if i + cons > end + 1:
                return False
            for j in range(cons):
                if self.grid[day][i + j].classtype != ClassType.NA or (i + j > NO_THEORY_FROM - 1 and not is_practical) or (day == 5 and i + j > BREAK_FROM - 1 and is_practical):
                    i += j + 1
                    flag = False
                    break
            if flag and batches != None and rooms != None:
                for j in range(cons):
                    no_of_sim_classes = 0
                    for batch in batches:
                        no_of_sim_classes += int(batch.grid[day][i + j].classtype != ClassType.NA)
                    if no_of_sim_classes >= len(rooms):
                        i += j + 1
                        flag = False
                        break
            if flag:
                for j in range(cons):
                    self.grid[day][i + j].classtype = classtype
                    self.grid[day][i + j].is_practical = is_practical
                    self.grid[day][i + j].cont = j != cons - 1
                if is_practical:
                    for pr in self.prs[classtype.name.lower()]:
                        if pr["cons"] == cons:
                            pr["freq"] -= 1
                            break
                else:
                    self.ths[classtype.name.lower()] -= 1
                return True
        return False

    def allocate(self, day, classtype, is_practical=False, cons=1, start=0, end=len(TIME_SLOTS)-1, batches=None, rooms=None):
        if start > end:
            return False
        if start < 0:
            start = 0
        if end >= len(TIME_SLOTS):
            end = len(TIME_SLOTS) - 1

        if start < BREAK_FROM:
            success = self.__allocate(day, classtype, is_practical, cons, start, end=BREAK_FROM-1 if end >= BREAK_FROM-1 else end, batches=batches, rooms=rooms)
            if not success and end > BREAK_FROM-1:
                return self.__allocate(day, classtype, is_practical, cons, start=BREAK_FROM, end=end, batches=batches, rooms=rooms)
            return success
        else:
            return self.__allocate(day, classtype, is_practical, cons, start=start, end=end, batches=batches, rooms=rooms)

    def rem_periods(self, classtype, is_practical=False):
        if is_practical:
            return sum([pr["cons"] * pr["freq"] for pr in self.prs[classtype.name.lower()]])
        return self.ths[classtype.name.lower()]

    def rem_classes(self, classtype, is_practical=False):
        if is_practical:
            return sum([pr["freq"] for pr in self.prs[classtype.name.lower()]])
        return self.ths[classtype.name.lower()]


def custom_distribute(batches, classtype1, classtype2):
    for day in random.sample(range(len(DAYS)), len(DAYS)):
        classtype1, classtype2 = classtype2, classtype1
        for batch in batches:
            if day == batch.offday:
                continue
            for classtype, (start, end) in zip([classtype1, classtype2], [(0, BREAK_FROM-1), (BREAK_FROM, len(TIME_SLOTS)-1)]):
                if (ths := batch.rem_classes(classtype)) + (prs := batch.rem_classes(classtype, is_practical=True)) > 0:
                    min_length = min(ths, prs)
                    if min_length > 0:
                        batch.allocate(day, classtype, start=start, end=end)
                        for pr in batch.prs[classtype.name.lower()]:
                            if pr["freq"] > 0:
                                cons = pr["cons"]
                                break
                        success = batch.allocate(day, classtype, is_practical=True, start=start, end=end, cons=cons)
                        if not success and batch.rem_classes(classtype) > 0:
                            batch.allocate(day, classtype, start=start, end=end)
                    elif batch.rem_classes(classtype) > 0:
                        batch.allocate(day, classtype, start=start, end=end)
                        if batch.rem_classes(classtype) > 0:
                            batch.allocate(day, classtype, start=start, end=end)
                    else:
                        for pr in batch.prs[classtype.name.lower()]:
                            if pr["freq"] > 0:
                                cons = pr["cons"]
                                break
                        batch.allocate(day, classtype, is_practical=True, start=start, end=end, cons=cons)
                        if batch.rem_classes(classtype, is_practical=True):
                            for pr in batch.prs[classtype.name.lower()]:
                                if pr["freq"] > 0:
                                    cons = pr["cons"]
                                    break
                            batch.allocate(day, classtype, is_practical=True, start=start, end=end, cons=cons)


def distribute(batches, rooms, classtype):
    # practical allocation
    limit1 = (0, BREAK_FROM-1)
    limit2 = (BREAK_FROM, len(TIME_SLOTS)-1)

    if classtype.name.lower() in batches[0].prs:
        for batch in batches:
            no_of_days = len(DAYS)
            days = random.sample(range(no_of_days), no_of_days)

            if batch.offday != None:
                days.remove(batch.offday)
                no_of_days -= 1

            while batch.rem_classes(classtype, is_practical=True) > 0 and len(days) > 0:
                for day in days:
                    if batch.rem_classes(classtype, is_practical=True) > 0:
                        for pr in batch.prs[classtype.name.lower()]:
                            if pr["freq"] > 0:
                                cons = pr["cons"]
                                break
                        success = batch.allocate(day, classtype, is_practical=True, cons=cons, start=limit1[0], end=limit1[1])
                        if not success:
                            success = batch.allocate(day, classtype, is_practical=True, cons=cons, start=limit2[0], end=limit2[1])
                        if not success:
                            days.remove(day)

                    limit1, limit2 = limit2, limit1

    # theory allocation
    if classtype.name.lower() in batches[0].ths:
        for batch in batches:
            no_of_days = len(DAYS)
            days = random.sample(range(no_of_days), no_of_days)

            if batch.offday != None:
                days.remove(batch.offday)
                no_of_days -= 1

            while batch.rem_classes(classtype) > 0 and len(days) > 0:
                for day in days:
                    if batch.rem_classes(classtype) > 0:
                        success = batch.allocate(day, classtype, batches=batches, rooms=rooms)
                        if not success:
                            success = batch.allocate(day, classtype, batches=batches, rooms=rooms)
                        if not success:
                            days.remove(day)


def allot_rooms(batches, homes, rooms, floors, classtype, is_practical=False):
    no_of_unallotted_periods = 0
    no_of_resolved_allotments = 0

    custom_stderr = ""

    for batch in batches:
        home_rooms = None

        for i in range(len(DAYS)):
            for j in range(len(TIME_SLOTS)):

                if batch.grid[i][j].room_id != None or batch.grid[i][j].classtype != classtype or batch.grid[i][j].is_practical != is_practical:
                    continue

                best_fit_room = None
                best_fit_diff = None

                cons = 1

                for k in range(j, len(TIME_SLOTS)):
                    if not batch.grid[i][k].cont or batch.grid[i][k].classtype != classtype or batch.grid[i][k].is_practical != is_practical:
                        break
                    cons += 1

                allotted_to_batches = set(b.grid[i][j + k].room_id for k in range(cons) for b in batches if b.grid[i][j + k].room_id != None)

                if home_rooms == None:
                    home_rooms = list(filter(lambda x: x.room_id in homes[batch.dept], rooms))

                for room in home_rooms:
                    if not room.room_id in allotted_to_batches and room.capacity >= batch.strength:
                        diff = room.capacity - batch.strength

                        if best_fit_room == None or diff < best_fit_diff:
                            best_fit_room = room
                            best_fit_diff = diff

                if best_fit_room != None:
                    for k in range(cons):
                        batch.grid[i][j + k].room_id = best_fit_room.room_id
                else:
                    for k in range(cons):
                        custom_stderr += f"warning: (pass 1) failed to allot a room for {batch.dept} {batch.prgm} {batch.sem} on {DAYS[i]} at {TIME_SLOTS[j + k]}\n"

    for batch in batches:
        home_rooms = None

        for i in range(len(DAYS)):
            for j in range(len(TIME_SLOTS)):

                if batch.grid[i][j].room_id != None or batch.grid[i][j].classtype != classtype or batch.grid[i][j].is_practical != is_practical:
                    continue

                best_fit_room = None
                best_fit_diff = None

                cons = 1

                for k in range(j, len(TIME_SLOTS)):
                    if not batch.grid[i][k].cont or batch.grid[i][k].classtype != classtype or batch.grid[i][k].is_practical != is_practical:
                        break
                    cons += 1

                allotted_to_batches = set(b.grid[i][j + k].room_id for k in range(cons) for b in batches if b.grid[i][j + k].room_id != None)

                if home_rooms == None:
                    home_rooms = list(filter(lambda x: x.room_id in homes[batch.dept], rooms))

                up_exists = j != 0 and batch.grid[i][j - 1].room_id != None
                down_exists = j != len(TIME_SLOTS) - 1 and batch.grid[i][j + 1].room_id != None

                if up_exists or down_exists:
                    floor_nos = set()

                    if up_exists:
                        for room in rooms:
                            if room.room_id == batch.grid[i][j - 1].room_id:
                                floor_nos.add(room.floor)
                                break

                    if down_exists:
                        for room in rooms:
                            if room.room_id == batch.grid[i][j + 1].room_id:
                                floor_nos.add(room.floor)
                                break

                    closest_floors = sorted(floors, key=lambda x: abs(x - sum(floor_nos) / len(floor_nos)))

                    for closest_floor in closest_floors:
                        best_fit_room = None
                        best_fit_diff = None

                        for room in filter(lambda x: x.floor == closest_floor, rooms):
                            if not room.room_id in allotted_to_batches and room.capacity >= batch.strength:
                                diff = room.capacity - batch.strength

                                if best_fit_room == None or diff < best_fit_diff:
                                    best_fit_room = room
                                    best_fit_diff = diff

                        if best_fit_room != None:
                            break
                else:
                    for room in rooms:
                        if not room.room_id in allotted_to_batches and room.capacity >= batch.strength:
                            diff = room.capacity - batch.strength

                            if best_fit_room == None or diff < best_fit_diff:
                                best_fit_room = room
                                best_fit_diff = diff

                if best_fit_room != None:
                    for k in range(cons):
                        batch.grid[i][j + k].room_id = best_fit_room.room_id
                    no_of_resolved_allotments += cons
                else:
                    for k in range(cons):
                        custom_stderr += f"error: (pass 2) failed to allot a room for {batch.dept} {batch.prgm} {batch.sem} on {DAYS[i]} at {TIME_SLOTS[j + k]}\n"

                    no_of_unallotted_periods += cons

    no_of_resolved_allotments -= no_of_unallotted_periods

    print(custom_stderr, end='')
    print("no of resolved allotments:", no_of_resolved_allotments)
    print("no of unallotted periods:", no_of_unallotted_periods)

    if no_of_unallotted_periods > 0:
        raise Exception


def init_offdays(batches):
    for batch in batches:
        offday_feasibility = batch.rem_periods(ClassType.MAJOR) + batch.rem_periods(ClassType.MAJOR, is_practical=True) + batch.rem_periods(ClassType.MINOR) + batch.rem_periods(ClassType.MINOR, is_practical=True) + batch.rem_periods(ClassType.MDS) + batch.rem_periods(ClassType.MDS, is_practical=True) + batch.rem_periods(ClassType.ENVS) <= len(TIME_SLOTS) * (len(DAYS) - 1)

        if offday_feasibility and batch.hasoffday:
            batch.offday = random.randint(0, len(DAYS) - 1)


def reorder_batches(depts_json, batches):
    ordered_batches = []
    for dept in depts_json["depts"]:
        for prgm in dept["prgms"]:
            for sem in prgm["sems"]:
                for batch in batches:
                    if dept["name"] == batch.dept and prgm["type"] == batch.prgm and sem["no"] == batch.sem:
                        ordered_batches.append(batch)
                        break
    return ordered_batches


def to_dataframe(batches):
    for batch in batches:
        df = pd.DataFrame(batch.grid, index=DAYS, columns=TIME_SLOTS)
        df = df.transpose().unstack().reset_index()
        df.columns = ["Day", "Time", f"{batch.dept} {batch.prgm} Sem {batch.sem}"]
        merged_df = pd.merge(merged_df, df, on=['Day', 'Time']) if batch != batches[0] else df

    return merged_df


def populate(depts_json, rooms_json, seed):
    random.seed(seed)

    batches = []
    homes = {}

    for dept in random.sample(depts_json["depts"], len(depts_json["depts"])):
        rooms = [room["room_id"] for room in dept["homes"]]
        homes[dept["name"]] = rooms

        for prgm in random.sample(dept["prgms"], len(dept["prgms"])):
            for sem in random.sample(prgm["sems"], len(prgm["sems"])):
                for key, value in sem["prs"].items():
                    sem["prs"][key] = sorted(value, key=lambda x: x["cons"])
                batches.append(Batch(dept["name"], prgm["type"], sem["no"], sem["hasoffday"], sem["strength"], sem["ths"], sem["prs"]))

    rooms = []
    floors = set()

    for room in rooms_json["rooms"]:
        rooms.append(Room(room["room_id"], room["capacity"], room["hasAC"], room["hasAV"], room["floor"]))
        floors.add(room["floor"])

    init_offdays(batches)

    custom_distribute(batches, ClassType.MINOR, ClassType.MDS)

    distribute(batches, rooms, ClassType.MAJOR)
    distribute(batches, rooms, ClassType.ENVS)

    allot_rooms(batches, homes, rooms, floors, ClassType.MAJOR)
    allot_rooms(batches, homes, rooms, floors, ClassType.MINOR)
    allot_rooms(batches, homes, rooms, floors, ClassType.MDS)
    allot_rooms(batches, homes, rooms, floors, ClassType.ENVS)

    ordered_batches = reorder_batches(depts_json, batches)
    return ordered_batches


def generate(depts_json, rooms_json, seed):
    i = 0
    while True:
        try:
            print("current seed:", seed)
            batches = populate(copy.deepcopy(depts_json), copy.deepcopy(rooms_json), seed)
            break
        except:
            print("error: failed to generate!")
            i += 1
            seed += (i * i + i) >> 1
            # seed += 1

    print(f"failed {i} times! generated with the seed value {seed}")
    return to_dataframe(batches)


def export(depts_json, rooms_json, output_path, seed):
    batches = generate(depts_json, rooms_json, seed)
    batches.to_csv(output_path, index=False)
    print(f"Saved {output_path}")
