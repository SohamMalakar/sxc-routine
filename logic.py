import random
import pandas as pd
import copy

from enum import Enum

DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
TIME_SLOTS = ("10:10-11:00", "11:05-11:55", "12:00-12:45", "1:15-2:05", "2:10-3:00", "3:05-3:55", "4:00-4:45")
BREAK_FROM = 3
NO_THEORY_FROM = 6
MIN_CLASSES = 2

ENVS_ROOM_ID = "24"

ClassType = Enum("ClassType", ['NA', 'MAJOR', 'MINOR', 'MDS', 'ENVS'])


class Slot:
    def __init__(self):
        self.room_id = None
        self.classtype = ClassType.NA
        self.is_practical = False
        self.next = False

    def __str__(self):
        if self.classtype != ClassType.NA:
            return f"{self.classtype.name}({self.is_practical and 'P' or 'T'}){'{Next}' * self.next} {str(self.room_id) * bool(self.room_id)}"
        return "X"


class Room:
    def __init__(self, room_id, capacity, hasAC, hasAV, floor):
        self.room_id = room_id
        self.capacity = capacity
        self.hasAC = hasAC
        self.hasAV = hasAV
        self.floor = floor


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
                    self.grid[day][i + j].next = j != cons - 1
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
            if not success and end > BREAK_FROM - 1:
                return self.__allocate(day, classtype, is_practical, cons, start=BREAK_FROM, end=end, batches=batches, rooms=rooms)
            return success
        else:
            return self.__allocate(day, classtype, is_practical, cons, start=start, end=end, batches=batches, rooms=rooms)

    def deallocate(self, day, index=0, ct=None):
        if index >= len(TIME_SLOTS):
            return None

        i = index
        cons = 0

        while True:
            ref = self.grid[day][i]

            classtype = ref.classtype

            if classtype == ClassType.NA:
                return i + 1

            if ct != None and ct != classtype:
                return None

            is_practical = ref.is_practical
            cons += 1

            ref.room_id = None
            ref.classtype = ClassType.NA
            ref.is_practical = False
            i += 1

            if not ref.next:
                break

            ref.next = False

        if is_practical:
            found = False

            for pr in self.prs[classtype.name.lower()]:
                if pr["cons"] == cons:
                    pr["freq"] += 1
                    found = True
                    break

            if not found:
                self.prs[classtype.name.lower()].append({"cons": cons, "freq": 1})
        else:
            self.ths[classtype.name.lower()] += 1
        
        return i

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


def distribute(batches, rooms, classtype, iteration=1):
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

            if batch.dept == "English" and batch.sem == 2 and batch.prgm == "U.G." and iteration == MIN_CLASSES:
                pass

            while batch.rem_classes(classtype, is_practical=True) > 0 and len(days) > 0:
                if batch.rem_classes(classtype) < MIN_CLASSES and iteration != 1:
                    days = sorted(days, key=(lambda x: sum([slot.classtype == ClassType.NA for slot in batch.grid[x]])))
                i = 0
                while i < len(days):
                    for _ in range(iteration):
                        if batch.rem_classes(classtype, is_practical=True) > 0:
                            for pr in batch.prs[classtype.name.lower()]:
                                if pr["freq"] > 0:
                                    cons = pr["cons"]
                                    break
                            success = batch.allocate(days[i], classtype, is_practical=True, cons=cons, start=limit1[0], end=limit1[1])
                            if not success:
                                success = batch.allocate(days[i], classtype, is_practical=True, cons=cons, start=limit2[0], end=limit2[1])
                            if not success:
                                days.remove(days[i])
                                i -= 1
                                break
                    i += 1

                    limit1, limit2 = limit2, limit1

    # theory allocation
    if classtype.name.lower() in batches[0].ths:
        for batch in batches:
            no_of_days = len(DAYS)
            days = random.sample(range(no_of_days), no_of_days)

            if batch.offday != None:
                days.remove(batch.offday)
                no_of_days -= 1

            if batch.dept == "English" and batch.sem == 2 and batch.prgm == "U.G." and iteration == MIN_CLASSES:
                pass

            while batch.rem_classes(classtype) > 0 and len(days) > 0:
                if batch.rem_classes(classtype) < MIN_CLASSES and iteration != 1:
                    days = sorted(days, key=(lambda x: sum([slot.classtype == ClassType.NA for slot in batch.grid[x]])))
                i = 0
                while i < len(days):
                    for _ in range(iteration):
                        if batch.rem_classes(classtype) > 0:
                            success = batch.allocate(days[i], classtype, batches=batches, rooms=rooms)
                            if not success:
                                days.remove(days[i])
                                i -= 1
                                break
                    i += 1


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
                    if not batch.grid[i][k].next or batch.grid[i][k].classtype != classtype or batch.grid[i][k].is_practical != is_practical:
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
                    if not batch.grid[i][k].next or batch.grid[i][k].classtype != classtype or batch.grid[i][k].is_practical != is_practical:
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


def init_offdays(batches, revoke=False):
    for batch in batches:
        rem_periods = batch.rem_periods(ClassType.MAJOR) + batch.rem_periods(ClassType.MAJOR, is_practical=True) + batch.rem_periods(ClassType.MINOR) + batch.rem_periods(ClassType.MINOR, is_practical=True) + batch.rem_periods(ClassType.MDS) + batch.rem_periods(ClassType.MDS, is_practical=True) + batch.rem_periods(ClassType.ENVS)

        if revoke and rem_periods > 0:
            batch.offday = None
            continue

        offday_feasibility = rem_periods <= (len(TIME_SLOTS) - 1) * (len(DAYS) - 1)

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


def allot_envs_room(batches):
    for batch in batches:
        for day in batch.grid:
            for slot in day:
                if slot.classtype == ClassType.ENVS:
                    slot.room_id = ENVS_ROOM_ID


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

    for batch in batches:
        for day in range(len(DAYS)):
            count = 0
            for slot in batch.grid[day]:
                if slot.classtype != ClassType.NA:
                    count += 1
            if count >= 1 and count < MIN_CLASSES:
                r = 0
                while r != None:
                    r = batch.deallocate(day, r, ClassType.MAJOR)

                r = 0
                while r != None:
                    r = batch.deallocate(day, r, ClassType.ENVS)

    init_offdays(batches, revoke=True)

    distribute(batches, rooms, ClassType.MAJOR, MIN_CLASSES)
    distribute(batches, rooms, ClassType.ENVS, MIN_CLASSES)

    for batch in batches:
        for day in range(len(DAYS)):
            count = 0
            for slot in batch.grid[day]:
                if slot.classtype != ClassType.NA:
                    count += 1
            if count >= 1 and count < MIN_CLASSES:
                raise Exception

    # if it fails to distribute the slots, it will throw an exception
    if sum([batch.rem_periods(ClassType.MAJOR) + batch.rem_periods(ClassType.MAJOR, is_practical=True) + batch.rem_periods(ClassType.MINOR) + batch.rem_periods(ClassType.MINOR, is_practical=True) + batch.rem_periods(ClassType.MDS) + batch.rem_periods(ClassType.MDS, is_practical=True) + batch.rem_periods(ClassType.ENVS) for batch in batches]) > 0:
        raise Exception

    # exclude the envs room from the rooms list
    rooms_wo_envs = list(filter(lambda room: room.room_id != ENVS_ROOM_ID, rooms))

    allot_rooms(batches, homes, rooms_wo_envs, floors, ClassType.MAJOR)
    allot_rooms(batches, homes, rooms_wo_envs, floors, ClassType.MINOR)
    allot_rooms(batches, homes, rooms_wo_envs, floors, ClassType.MDS)

    # one specific room will be alloted for all envs classes
    allot_envs_room(batches)

    ordered_batches = reorder_batches(depts_json, batches)
    return ordered_batches


def generate(depts_json, rooms_json, seed, iteration=1000):
    batches = None

    for i in range(iteration):
        try:
            print("current seed:", seed)
            batches = populate(copy.deepcopy(depts_json), copy.deepcopy(rooms_json), seed)
            break
        except:
            print("error: failed to generate!")
            seed += (i * i + i) >> 1

    # if it fails to generate the routine even after {iteration} no of iterations, it will throw an exception
    if batches == None:
        raise Exception

    print(f"failed {i} times! generated with the seed value {seed}")
    return to_dataframe(batches)