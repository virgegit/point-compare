"""Core comparison logic for point-compare."""

import pandas as pd
import numpy as np

from .schema import LOGICAL_FIELDS


def coord_key(row, tol, ijk_tol, use_ijk):
    """Tolerance-quantized coordinate key for bucketing."""
    def rnd(v, t):
        try:
            return int(round(float(v) / t))
        except (TypeError, ValueError):
            return None
    k = (rnd(row["X"], tol), rnd(row["Y"], tol), rnd(row["Z"], tol))
    if use_ijk:
        k += (rnd(row["I"], ijk_tol), rnd(row["J"], ijk_tol), rnd(row["K"], ijk_tol))
    return k


def safe_delta(v1, v2):
    """Safe difference with NaN handling."""
    try:
        return round(float(v2) - float(v1), 4)
    except (TypeError, ValueError):
        return np.nan


def fmt_val(v):
    """Format value for display."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "–"
    if isinstance(v, float):
        return f"{v:.4f}"
    return str(v)


def compare(df_orig, df_new, tol, ijk_tol, use_ijk):
    """Main comparison algorithm."""
    name_to_new = {r["Name"]: i for i, r in df_new.iterrows()}
    coord_to_new = {}
    for i, r in df_new.iterrows():
        k = coord_key(r, tol, ijk_tol, use_ijk)
        coord_to_new.setdefault(k, []).append(i)

    matched_new = set()
    results = []

    for _, r1 in df_orig.iterrows():
        n1 = r1["Name"]
        ck1 = coord_key(r1, tol, ijk_tol, use_ijk)
        by_name = n1 in name_to_new
        by_coord = ck1 in coord_to_new

        row = {f"ORIG_{f}": r1[f] for f in LOGICAL_FIELDS}
        row.update({f"NEW_{f}": np.nan for f in LOGICAL_FIELDS})
        row.update({"NAME_diff": "", "X_diff": np.nan, "Y_diff": np.nan, "Z_diff": np.nan,
                    "I_diff": np.nan, "J_diff": np.nan, "K_diff": np.nan, "DIFF_Fields": ""})

        if by_name:
            i2 = name_to_new[n1]
            r2 = df_new.loc[i2]
            matched_new.add(i2)
            row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
            ck2 = coord_key(r2, tol, ijk_tol, use_ijk)
            if ck1 == ck2:
                row["Status"] = "MATCH"
            else:
                row["Status"] = "COORD_CHANGED"
                diffs = []
                for f in ["X","Y","Z","I","J","K"]:
                    d = safe_delta(r1[f], r2[f])
                    row[f"{f}_diff"] = d
                    if not (isinstance(d, float) and np.isnan(d)) and abs(d) > 1e-9:
                        diffs.append(f)
                row["DIFF_Fields"] = ", ".join(diffs)

        elif by_coord:
            found = False
            for i2 in coord_to_new[ck1]:
                if i2 not in matched_new:
                    r2 = df_new.loc[i2]
                    matched_new.add(i2)
                    row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
                    row["Status"] = "NAME_CHANGED"
                    row["NAME_diff"] = f"{n1}  →  {r2['Name']}"
                    row["DIFF_Fields"] = "Name"
                    found = True
                    break
            if not found:
                row["Status"] = "DELETED"
        else:
            row["Status"] = "DELETED"

        results.append(row)

    for i2, r2 in df_new.iterrows():
        if i2 not in matched_new:
            row = {"Status": "ADDED"}
            row.update({f"ORIG_{f}": np.nan for f in LOGICAL_FIELDS})
            row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
            row.update({"NAME_diff": "", "DIFF_Fields": "", "X_diff": np.nan, "Y_diff": np.nan,
                        "Z_diff": np.nan, "I_diff": np.nan, "J_diff": np.nan, "K_diff": np.nan})
            results.append(row)

    return pd.DataFrame(results)


def build_close_points_report(df_orig, df_new, distance_tol, df_all=None):
    """Find ALL original/new point pairs within distance_tol with status info."""
    orig_xyz = df_orig[["X", "Y", "Z"]].to_numpy(dtype=float)
    new_xyz = df_new[["X", "Y", "Z"]].to_numpy(dtype=float)
    results = []

    for pos_idx, (_, orig_row) in enumerate(df_orig.iterrows()):
        orig_point = orig_xyz[pos_idx]
        if np.isnan(orig_point).any():
            continue
        deltas = new_xyz - orig_point
        axis_mask = ~np.isnan(deltas).any(axis=1)
        if not axis_mask.any():
            continue
        candidate_deltas = deltas[axis_mask]
        in_box = np.all(np.abs(candidate_deltas) <= distance_tol, axis=1)
        if not in_box.any():
            continue
        candidate_positions = np.where(axis_mask)[0][in_box]
        for new_pos, delta in zip(candidate_positions, candidate_deltas[in_box]):
            distance = float(np.linalg.norm(delta))
            if distance <= 1e-12 or distance > distance_tol + 1e-12:
                continue
            new_row = df_new.iloc[new_pos]

            status = "REPLACED?"
            if df_all is not None:
                try:
                    row_status = df_all.iloc[pos_idx]["Status"]
                    status = row_status
                except:
                    pass

            name_diff = f"{orig_row['Name']} → {new_row['Name']}" if orig_row["Name"] != new_row["Name"] else ""

            results.append({
                "Status": status,
                "ORIG_Name": orig_row["Name"],
                "ORIG_X": orig_row["X"],
                "ORIG_Y": orig_row["Y"],
                "ORIG_Z": orig_row["Z"],
                "NEW_Name": new_row["Name"],
                "NEW_X": new_row["X"],
                "NEW_Y": new_row["Y"],
                "NEW_Z": new_row["Z"],
                "NAME_diff": name_diff,
                "dX": safe_delta(orig_row["X"], new_row["X"]),
                "dY": safe_delta(orig_row["Y"], new_row["Y"]),
                "dZ": safe_delta(orig_row["Z"], new_row["Z"]),
                "Distance_3D": round(distance, 4)
            })

    if not results:
        return pd.DataFrame(columns=["Status", "ORIG_Name", "ORIG_X", "ORIG_Y", "ORIG_Z", "NEW_Name", "NEW_X",
                                     "NEW_Y", "NEW_Z", "NAME_diff", "dX", "dY", "dZ", "Distance_3D"])
    return pd.DataFrame(results).sort_values(by=["Distance_3D", "ORIG_Name", "NEW_Name"],
                                             kind="stable").reset_index(drop=True)


def build_nearest_points_report(df_orig, df_new):
    """For each original point, find the nearest new point."""
    orig_xyz = df_orig[["X", "Y", "Z"]].to_numpy(dtype=float)
    new_xyz = df_new[["X", "Y", "Z"]].to_numpy(dtype=float)
    results = []

    for pos_idx, (_, orig_row) in enumerate(df_orig.iterrows()):
        orig_point = orig_xyz[pos_idx]
        if np.isnan(orig_point).any():
            continue
        deltas = new_xyz - orig_point
        valid_mask = ~np.isnan(deltas).any(axis=1)
        if not valid_mask.any():
            continue
        valid_deltas = deltas[valid_mask]
        distances = np.linalg.norm(valid_deltas, axis=1)
        if len(distances) == 0:
            continue
        nearest_pos = int(np.argmin(distances))
        all_valid_positions = np.where(valid_mask)[0]
        new_pos = int(all_valid_positions[nearest_pos])
        distance = float(distances[nearest_pos])
        new_row = df_new.iloc[new_pos]
        results.append({"ORIG_Name": orig_row["Name"], "ORIG_X": orig_row["X"], "ORIG_Y": orig_row["Y"],
                       "ORIG_Z": orig_row["Z"], "NEW_Name": new_row["Name"], "NEW_X": new_row["X"],
                       "NEW_Y": new_row["Y"], "NEW_Z": new_row["Z"],
                       "dX": safe_delta(orig_row["X"], new_row["X"]),
                       "dY": safe_delta(orig_row["Y"], new_row["Y"]),
                       "dZ": safe_delta(orig_row["Z"], new_row["Z"]),
                       "Distance_3D": round(distance, 4),
                       "Same_Name": "Yes" if str(orig_row["Name"]) == str(new_row["Name"]) else "No"})

    if not results:
        return pd.DataFrame(columns=["ORIG_Name", "ORIG_X", "ORIG_Y", "ORIG_Z", "NEW_Name", "NEW_X",
                                     "NEW_Y", "NEW_Z", "dX", "dY", "dZ", "Distance_3D", "Same_Name"])
    return pd.DataFrame(results).sort_values(by=["Distance_3D", "ORIG_Name", "NEW_Name"],
                                             kind="stable").reset_index(drop=True)


def link_moved_pairs(df_all, df_orig, df_new, close_points_tol):
    """Relabel DELETED+ADDED pairs within tolerance as REPLACED?.

    Uses ORIG_*/NEW_* values from df_all directly (avoids index-mapping bugs).
    For each DELETED point, finds the closest ADDED point within tolerance.
    Supports many-to-one: multiple DELETED rows can share the same ADDED match.
    """
    deleted_indices = list(df_all[df_all["Status"] == "DELETED"].index)
    added_indices = list(df_all[df_all["Status"] == "ADDED"].index)
    if not deleted_indices or not added_indices:
        return df_all, []

    added_points = []
    for add_label in added_indices:
        add_row = df_all.loc[add_label]
        try:
            pt = np.array([float(add_row["NEW_X"]), float(add_row["NEW_Y"]), float(add_row["NEW_Z"])])
            if not np.isnan(pt).any():
                added_points.append((add_label, pt, add_row["NEW_Name"]))
        except (TypeError, ValueError):
            continue

    if not added_points:
        return df_all, []

    df_all = df_all.copy()
    moved_pairs = []
    used_added_labels = set()

    for del_label in deleted_indices:
        del_row = df_all.loc[del_label]
        try:
            orig_point = np.array([float(del_row["ORIG_X"]), float(del_row["ORIG_Y"]), float(del_row["ORIG_Z"])])
        except (TypeError, ValueError):
            continue
        if np.isnan(orig_point).any():
            continue

        best_add_label, best_dist, best_new_name, best_new_point = None, float("inf"), None, None
        for add_label, new_point, new_name in added_points:
            dist = float(np.linalg.norm(orig_point - new_point))
            if dist <= close_points_tol + 1e-12 and dist < best_dist:
                best_dist = dist
                best_add_label = add_label
                best_new_name = new_name
                best_new_point = new_point

        if best_add_label is not None:
            add_row = df_all.loc[best_add_label]
            df_all.at[del_label, "Status"] = "REPLACED?"
            for f in LOGICAL_FIELDS:
                df_all.at[del_label, f"NEW_{f}"] = add_row[f"NEW_{f}"]
            orig_name = del_row["ORIG_Name"]
            df_all.at[del_label, "NAME_diff"] = (
                f"{orig_name}  →  {best_new_name}" if orig_name != best_new_name else ""
            )
            for f in ["X", "Y", "Z", "I", "J", "K"]:
                df_all.at[del_label, f"{f}_diff"] = safe_delta(del_row[f"ORIG_{f}"], add_row[f"NEW_{f}"])
            diffs = []
            for f in ["X", "Y", "Z", "I", "J", "K"]:
                d = df_all.at[del_label, f"{f}_diff"]
                if not (isinstance(d, float) and np.isnan(d)) and abs(d) > 1e-9:
                    diffs.append(f)
            df_all.at[del_label, "DIFF_Fields"] = ", ".join(diffs)

            used_added_labels.add(best_add_label)
            moved_pairs.append({
                "ORIG_Name": del_row["ORIG_Name"],
                "ORIG_X": float(orig_point[0]), "ORIG_Y": float(orig_point[1]), "ORIG_Z": float(orig_point[2]),
                "NEW_Name": best_new_name,
                "NEW_X": float(best_new_point[0]), "NEW_Y": float(best_new_point[1]), "NEW_Z": float(best_new_point[2]),
                "dX": safe_delta(orig_point[0], best_new_point[0]),
                "dY": safe_delta(orig_point[1], best_new_point[1]),
                "dZ": safe_delta(orig_point[2], best_new_point[2]),
                "Distance_3D": round(best_dist, 4),
            })

    for add_label in used_added_labels:
        df_all.at[add_label, "Status"] = "REPLACED?"

    return df_all, moved_pairs
