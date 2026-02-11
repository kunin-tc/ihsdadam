"""QGraphicsView-based highway alignment visualization widget"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGraphicsView, QGraphicsScene,
    QGraphicsRectItem, QGraphicsLineItem, QGraphicsTextItem,
    QGraphicsPolygonItem, QGraphicsPathItem, QGraphicsSimpleTextItem,
)
from PySide6.QtCore import Qt, QRectF, QPointF, Signal, QTimer
from PySide6.QtGui import (
    QPen, QBrush, QColor, QFont, QPainterPath, QPolygonF,
    QPainter,
)

from ..theme import (
    LANE_ASPHALT, LANE_LEFT_TURN, SHOULDER_GRAY, CENTERLINE_YELLOW,
    PRIMARY, TEXT_PRIMARY, TEXT_SECONDARY, ACCENT_GREEN, DANGER, WARNING,
)


def _format_station(station_str):
    if not station_str or str(station_str).strip() == "":
        return str(station_str)
    try:
        val = float(station_str)
        if val < 100:
            return str(station_str)
        parts = str(val).split(".")
        integer = parts[0]
        decimal = parts[1] if len(parts) > 1 else "00"
        decimal = decimal.ljust(2, "0")
        if len(integer) >= 2:
            return f"{integer[:-2]}+{integer[-2:]}.{decimal}"
        return f"{integer}.{decimal}"
    except (ValueError, TypeError):
        return str(station_str)


class _ZoomableView(QGraphicsView):
    """QGraphicsView with wheel-zoom, click-drag pan, and auto-fit."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._auto_fit = True

    def wheelEvent(self, event):
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.scale(factor, factor)
        self._auto_fit = False
        event.accept()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Home, Qt.Key_0):
            self.fit_contents()
            event.accept()
        else:
            super().keyPressEvent(event)

    def fit_contents(self):
        """Fit scene contents into the viewport."""
        scene = self.scene()
        if scene and not scene.sceneRect().isEmpty():
            self.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)
            self._auto_fit = True

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._auto_fit:
            scene = self.scene()
            if scene and not scene.sceneRect().isEmpty():
                self.fitInView(scene.sceneRect(), Qt.KeepAspectRatio)


class HighwayCanvas(QWidget):
    """Full highway visualization with stacked panels drawn via QGraphicsScene."""

    item_hovered = Signal(str)   # tooltip text when hovering an element

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._view = _ZoomableView()
        self._view.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        self._view.setDragMode(QGraphicsView.ScrollHandDrag)
        self._view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self._view.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self._view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        layout.addWidget(self._view)

        self._scene = QGraphicsScene(self)
        self._view.setScene(self._scene)

        self._data = None

    def clear(self):
        self._scene.clear()
        self._data = None

    def set_data(self, data: dict):
        """Accept parsed highway data dict and redraw everything."""
        self._data = data
        self._draw()
        # Delay fitInView so the layout has settled before measuring
        QTimer.singleShot(0, self._view.fit_contents)

    # ── drawing ────────────────────────────────────────────────────────────

    def _draw(self):
        self._scene.clear()
        d = self._data
        if not d:
            return

        min_sta = d["min_sta"]
        max_sta = d["max_sta"]
        title = d.get("title", "")
        heading_sta = d.get("heading_sta", "0")
        heading_angle = d.get("heading_angle", "0")

        canvas_w = max(self._view.viewport().width(), 900)
        margin = 40
        pw = canvas_w - 2 * margin

        def sta_x(station):
            if max_sta == min_sta:
                return margin + pw / 2
            return margin + 10 + (station - min_sta) / (max_sta - min_sta) * (pw - 20)

        y = 0.0

        # ── header ─────────────────────────────────────────────────────
        hdr_h = 90
        self._scene.addRect(QRectF(0, y, canvas_w, hdr_h),
                            QPen(QColor("#cccccc")), QBrush(QColor("#f0f0f0")))
        self._add_text(20, y + 12, f"Highway Alignment: {title}",
                       bold=True, size=13, color=PRIMARY)
        length = max_sta - min_sta
        self._add_text(20, y + 40,
                       f"Station Range: {_format_station(str(min_sta))} to "
                       f"{_format_station(str(max_sta))} ({length:.2f} ft)", size=9)
        self._add_text(20, y + 58,
                       f"Heading: {heading_angle} deg at Sta {_format_station(str(heading_sta))}",
                       size=9)
        y += hdr_h + 15

        # ── lane plan ──────────────────────────────────────────────────
        lanes = d.get("lanes", [])
        shoulders = d.get("shoulders", [])
        connections = self._build_connections(d)
        if lanes or shoulders:
            y = self._draw_lane_panel(y, canvas_w, margin, pw, sta_x, lanes,
                                      shoulders, connections, min_sta, max_sta)
            y += 25

        # ── ramp locations ─────────────────────────────────────────────
        ramps = d.get("ramps", [])
        if ramps:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Ramp Locations", ramps, self._draw_ramp_markers, 80)
            y += 25

        # ── horizontal curves ──────────────────────────────────────────
        curves = d.get("curves", [])
        if curves:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Horizontal Alignment (Curves & Tangents)", curves,
                self._draw_curve_markers, 100)
            y += 25

        # ── traffic ────────────────────────────────────────────────────
        traffic = d.get("traffic", [])
        if traffic:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Annual Average Daily Traffic (AADT)", traffic,
                self._draw_traffic_bars, 100)
            y += 25

        # ── median ─────────────────────────────────────────────────────
        median = d.get("median", [])
        if median:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Median Width", median, self._draw_median_bars, 70)
            y += 25

        # ── speed ──────────────────────────────────────────────────────
        speed = d.get("speed", [])
        if speed:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Posted Speed Limit", speed, self._draw_speed_bars, 60)
            y += 25

        # ── functional class ───────────────────────────────────────────
        func_class = d.get("func_class", [])
        if func_class:
            y = self._draw_simple_panel(
                y, canvas_w, margin, pw, sta_x, min_sta, max_sta,
                "Functional Classification", func_class,
                self._draw_func_class_bars, 60)
            y += 20

        self._scene.setSceneRect(0, 0, canvas_w, y + 20)

    # ── lane panel (most complex) ──────────────────────────────────────

    def _draw_lane_panel(self, y_start, cw, mx, pw, sta_x, lanes, shoulders,
                         connections, min_sta, max_sta):
        self._add_text(mx, y_start, "Lane Configuration (Plan View)",
                       bold=True, size=10, color=PRIMARY)
        y_start += 22

        lane_h = 18
        sh_h = 12

        right_lanes = sorted([l for l in lanes if l["side"] in ("right", "both")],
                             key=lambda x: x["priority"])
        left_lanes = sorted([l for l in lanes if l["side"] in ("left", "both")],
                            key=lambda x: x["priority"])

        r_prios = sorted(set(l["priority"] for l in right_lanes))
        l_prios = sorted(set(l["priority"] for l in left_lanes))
        max_r = len(r_prios)
        max_l = len(l_prios)

        has_ro = any(s.get("position", "outside") == "outside" and s["side"] in ("right", "both") for s in shoulders)
        has_lo = any(s.get("position", "outside") == "outside" and s["side"] in ("left", "both") for s in shoulders)
        has_ri = any(s.get("position", "outside") == "inside" and s["side"] in ("right", "both") for s in shoulders)
        has_li = any(s.get("position", "outside") == "inside" and s["side"] in ("left", "both") for s in shoulders)

        total_above = max_l * lane_h + (sh_h if has_lo else 0) + (sh_h if has_li else 0)
        total_below = max_r * lane_h + (sh_h if has_ro else 0) + (sh_h if has_ri else 0)
        panel_h = max(120, total_above + total_below + 70)

        self._scene.addRect(QRectF(mx, y_start, pw, panel_h),
                            QPen(QColor("#666666"), 2), QBrush(QColor("#f5f5f5")))

        cl_y = y_start + (sh_h if has_lo else 0) + (sh_h if has_li else 0) + max_l * lane_h + 35

        # station breakpoints
        all_sta = sorted(set(l["begin"] for l in lanes) | set(l["end"] for l in lanes))

        def _slot(lane, station, side):
            if side == "right":
                active = [l for l in right_lanes if l["begin"] <= station <= l["end"]]
            else:
                active = [l for l in left_lanes if l["begin"] <= station <= l["end"]]
            ap = sorted(set(l["priority"] for l in active))
            return ap.index(lane["priority"]) if lane["priority"] in ap else 0

        def _count(station, side):
            if side == "right":
                active = [l for l in right_lanes if l["begin"] <= station <= l["end"]]
            else:
                active = [l for l in left_lanes if l["begin"] <= station <= l["end"]]
            return len(set(l["priority"] for l in active))

        def _has_inside(station, side):
            for s in shoulders:
                if s.get("position", "outside") == "inside":
                    if (s["side"] == side or s["side"] == "both") and s["begin"] <= station <= s["end"]:
                        return True
            return False

        # draw lanes
        for lane in lanes:
            color = LANE_LEFT_TURN if lane.get("lane_type") == "left_turn" else LANE_ASPHALT
            side = lane["side"]
            sides = ["left", "right"] if side == "both" else [side]
            lane_type = lane.get("lane_type", "through")
            width_str = f"\nWidth: {lane['width']} ft" if lane.get("width") else ""
            for ds in sides:
                seg_sta = sorted(set(s for s in all_sta if lane["begin"] <= s <= lane["end"]) | {lane["begin"], lane["end"]})
                for i in range(len(seg_sta) - 1):
                    ss, se = seg_sta[i], seg_sta[i + 1]
                    mid = (ss + se) / 2
                    slot = _slot(lane, mid, ds)
                    io = sh_h if _has_inside(mid, ds) else 0
                    x0, x1 = sta_x(ss), sta_x(se)
                    if ds == "right":
                        yt = cl_y + io + slot * lane_h
                    else:
                        yt = cl_y - io - (slot + 1) * lane_h
                    item = self._scene.addRect(
                        QRectF(x0, yt, x1 - x0, lane_h),
                        QPen(Qt.NoPen), QBrush(QColor(color)))
                    self._set_hover(item,
                        f"Lane: {lane_type}\nSide: {ds}"
                        f"\nSta {_format_station(str(ss))} \u2013 "
                        f"{_format_station(str(se))}{width_str}")

        # draw shoulders
        for sh in shoulders:
            pos = sh.get("position", "outside")
            side = sh["side"]
            sides = ["left", "right"] if side == "both" else [side]
            width_str = f"\nWidth: {sh['width']} ft" if sh.get("width") else ""
            for ds in sides:
                seg_sta = sorted(set(s for s in all_sta if sh["begin"] <= s <= sh["end"]) | {sh["begin"], sh["end"]})
                for i in range(len(seg_sta) - 1):
                    ss, se = seg_sta[i], seg_sta[i + 1]
                    mid = (ss + se) / 2
                    na = _count(mid, ds)
                    io = sh_h if _has_inside(mid, ds) else 0
                    x0, x1 = sta_x(ss), sta_x(se)
                    if pos == "outside":
                        if ds == "right":
                            yt = cl_y + io + na * lane_h
                        else:
                            yt = cl_y - io - na * lane_h - sh_h
                    else:
                        if ds == "right":
                            yt = cl_y
                        else:
                            yt = cl_y - sh_h
                    item = self._scene.addRect(
                        QRectF(x0, yt, x1 - x0, sh_h),
                        QPen(Qt.NoPen), QBrush(QColor(SHOULDER_GRAY)))
                    self._set_hover(item,
                        f"Shoulder ({pos})\nSide: {ds}"
                        f"\nSta {_format_station(str(ss))} \u2013 "
                        f"{_format_station(str(se))}{width_str}")

        # centerline
        pen_cl = QPen(QColor(CENTERLINE_YELLOW), 2, Qt.DashLine)
        pen_cl.setDashPattern([12, 6])
        self._scene.addLine(mx + 10, cl_y, mx + pw - 10, cl_y, pen_cl)

        # lane markings
        for i in range(len(all_sta) - 1):
            ss, se = all_sta[i], all_sta[i + 1]
            mid = (ss + se) / 2
            for side in ("right", "left"):
                if side == "right":
                    active = [l for l in right_lanes if l["begin"] <= mid <= l["end"]]
                else:
                    active = [l for l in left_lanes if l["begin"] <= mid <= l["end"]]
                ap = sorted(set(l["priority"] for l in active))
                io = sh_h if _has_inside(mid, side) else 0
                for j in range(len(ap) - 1):
                    if side == "right":
                        my = cl_y + io + (j + 1) * lane_h
                    else:
                        my = cl_y - io - (j + 1) * lane_h
                    pen_w = QPen(QColor("white"), 1, Qt.DashLine)
                    pen_w.setDashPattern([8, 4])
                    self._scene.addLine(sta_x(ss), my, sta_x(se), my, pen_w)

        # connection markers
        road_top = cl_y - (sh_h if has_li else 0) - max_l * lane_h - (sh_h if has_lo else 0)
        road_bot = cl_y + (sh_h if has_ri else 0) + max_r * lane_h + (sh_h if has_ro else 0)
        for conn in connections:
            cs = conn["station"]
            if cs < min_sta or cs > max_sta:
                continue
            x = sta_x(cs)
            lt, lb = road_top - 15, road_bot + 15
            ct = conn.get("type", "ramp")
            conn_name = conn.get("name", ct.title())
            if ct == "intersection":
                pen = QPen(QColor(ACCENT_GREEN), 3, Qt.DashLine)
                pen.setDashPattern([6, 3])
                line_item = self._scene.addLine(x, lt, x, lb, pen)
                self._set_hover(line_item,
                    f"Intersection: {conn_name}"
                    f"\nSta {_format_station(str(cs))}")
                poly = QPolygonF([QPointF(x, lt - 8), QPointF(x - 5, lt),
                                  QPointF(x, lt + 8), QPointF(x + 5, lt)])
                poly_item = self._scene.addPolygon(
                    poly, QPen(QColor("#047857")), QBrush(QColor(ACCENT_GREEN)))
                self._set_hover(poly_item,
                    f"Intersection: {conn_name}"
                    f"\nSta {_format_station(str(cs))}")
            else:
                rt = conn.get("ramp_type", "entrance")
                c = QColor(WARNING) if rt == "entrance" else QColor(DANGER)
                pen = QPen(c, 3, Qt.DashLine)
                pen.setDashPattern([6, 3])
                line_item = self._scene.addLine(x, lt, x, lb, pen)
                self._set_hover(line_item,
                    f"Ramp: {conn_name} ({rt})"
                    f"\nSta {_format_station(str(cs))}")
                if rt == "entrance":
                    tri = QPolygonF([QPointF(x, lt - 8), QPointF(x - 5, lt + 2),
                                     QPointF(x + 5, lt + 2)])
                else:
                    tri = QPolygonF([QPointF(x, lb + 8), QPointF(x - 5, lb - 2),
                                     QPointF(x + 5, lb - 2)])
                tri_item = self._scene.addPolygon(tri, QPen(c.darker(120)), QBrush(c))
                self._set_hover(tri_item,
                    f"Ramp: {conn_name} ({rt})"
                    f"\nSta {_format_station(str(cs))}")

        # station markers
        mk_y = y_start + panel_h - 20
        for i in range(5):
            frac = i / 4
            x = mx + 10 + (pw - 20) * frac
            sta = min_sta + (max_sta - min_sta) * frac
            self._scene.addLine(x, mk_y - 8, x, mk_y + 8,
                                QPen(QColor("#333333"), 2))
            self._add_text(x, mk_y + 10, _format_station(str(sta)),
                           size=9, bold=True, anchor="top-center")

        # direction labels
        self._add_text(mx + pw - 10, cl_y + 12, "Right side (forward)",
                       size=8, color=TEXT_SECONDARY, anchor="top-right")
        self._add_text(mx + pw - 10, cl_y - 20, "Left side (opposite)",
                       size=8, color=TEXT_SECONDARY, anchor="top-right")

        return y_start + panel_h

    # ── generic panel helpers ──────────────────────────────────────────

    def _draw_simple_panel(self, y_start, cw, mx, pw, sta_x, min_sta, max_sta,
                           title, data, draw_fn, panel_h):
        self._add_text(mx, y_start, title, bold=True, size=10, color=PRIMARY)
        y_start += 22
        self._scene.addRect(QRectF(mx, y_start, pw, panel_h),
                            QPen(QColor("#666666"), 2), QBrush(QColor("white")))
        draw_fn(y_start, mx, pw, sta_x, min_sta, max_sta, data, panel_h)
        return y_start + panel_h

    def _draw_ramp_markers(self, y0, mx, pw, sta_x, min_sta, max_sta, ramps, ph):
        bl = y0 + ph // 2
        self._scene.addLine(mx + 10, bl, mx + pw - 10, bl,
                            QPen(QColor("#666666"), 2))
        for r in ramps:
            x = sta_x(r["station"])
            rt = r.get("ramp_type", "entrance")
            c = QColor(ACCENT_GREEN) if rt == "entrance" else QColor(DANGER)
            if rt == "entrance":
                tri = QPolygonF([QPointF(x, bl - 15), QPointF(x - 6, bl - 3),
                                 QPointF(x + 6, bl - 3)])
            else:
                tri = QPolygonF([QPointF(x, bl + 15), QPointF(x - 6, bl + 3),
                                 QPointF(x + 6, bl + 3)])
            ramp_name = r.get("name", "Ramp")
            item = self._scene.addPolygon(tri, QPen(QColor(PRIMARY)), QBrush(c))
            self._set_hover(item,
                f"{ramp_name}\nType: {rt}"
                f"\nSta {_format_station(str(r['station']))}")
            ty = bl + (28 if rt == "exit" else -28)
            self._add_text(x, ty, ramp_name, size=7, anchor="top-center")

    def _draw_curve_markers(self, y0, mx, pw, sta_x, min_sta, max_sta, curves, ph):
        bl = y0 + ph // 2
        for curve in curves:
            x0 = sta_x(curve["begin"])
            x1 = sta_x(curve["end"])
            sta_range = (f"Sta {_format_station(str(curve['begin']))} \u2013 "
                         f"{_format_station(str(curve['end']))}")
            if curve["type"] == "tangent":
                item = self._scene.addLine(x0, bl, x1, bl,
                                           QPen(QColor("#3b82f6"), 4))
                self._set_hover(item, f"Tangent\n{sta_range}")
                self._add_text((x0 + x1) / 2, bl - 15, "Tangent",
                               size=7, color="#666666", anchor="top-center")
            else:
                direction = curve.get("direction", "left")
                radius = curve.get("radius", 0)
                y_off = -20 if direction == "left" else 20
                path = QPainterPath()
                path.moveTo(x0, bl)
                steps = 10
                for i in range(1, steps + 1):
                    t = i / steps
                    cx = x0 + t * (x1 - x0)
                    cy = bl + y_off * (4 * t * (1 - t))
                    path.lineTo(cx, cy)
                pen_c = QPen(QColor("#f59e0b"), 4)
                item = self._scene.addPath(path, pen_c)
                self._set_hover(item,
                    f"Curve ({direction})\nRadius: {radius:.0f} ft\n{sta_range}")
                ly = bl + y_off + (15 if direction == "right" else -15)
                self._add_text((x0 + x1) / 2, ly, f"R={radius:.0f}'",
                               size=7, color="#f59e0b", anchor="top-center")

    def _draw_traffic_bars(self, y0, mx, pw, sta_x, min_sta, max_sta, traffic, ph):
        if not traffic:
            return
        max_vol = max(t["volume"] for t in traffic) or 1
        for t in traffic:
            x0 = sta_x(t["begin"])
            x1 = sta_x(t["end"])
            bh = (t["volume"] / max_vol) * (ph - 40)
            yb = y0 + ph - 20 - bh
            item = self._scene.addRect(
                QRectF(x0, yb, x1 - x0, bh),
                QPen(QColor("#5b21b6")), QBrush(QColor("#8b5cf6")))
            self._set_hover(item,
                f"AADT: {t['volume']:,}"
                f"\nSta {_format_station(str(t['begin']))} \u2013 "
                f"{_format_station(str(t['end']))}")
            self._add_text((x0 + x1) / 2, yb - 3, f"{t['volume']:,}",
                           size=8, color="#5b21b6", anchor="bottom-center")

    def _draw_median_bars(self, y0, mx, pw, sta_x, min_sta, max_sta, medians, ph):
        bl = y0 + ph - 20
        for m in medians:
            x0, x1 = sta_x(m["begin"]), sta_x(m["end"])
            item = self._scene.addRect(
                QRectF(x0, y0 + 15, x1 - x0, bl - y0 - 15),
                QPen(QColor("#78716c")), QBrush(QColor("#a8a29e")))
            median_type = m.get("median_type", "")
            self._set_hover(item,
                f"Median: {m['width']:.0f} ft {median_type}"
                f"\nSta {_format_station(str(m['begin']))} \u2013 "
                f"{_format_station(str(m['end']))}")
            self._add_text((x0 + x1) / 2, y0 + 25,
                           f"{m['width']:.0f}' {median_type}",
                           size=8, color="white", anchor="top-center")

    def _draw_speed_bars(self, y0, mx, pw, sta_x, min_sta, max_sta, speeds, ph):
        for s in speeds:
            x0, x1 = sta_x(s["begin"]), sta_x(s["end"])
            item = self._scene.addRect(
                QRectF(x0, y0 + 10, x1 - x0, ph - 20),
                QPen(QColor("#f59e0b"), 2), QBrush(QColor("#fef3c7")))
            self._set_hover(item,
                f"Speed Limit: {s['speed']} mph"
                f"\nSta {_format_station(str(s['begin']))} \u2013 "
                f"{_format_station(str(s['end']))}")
            self._add_text((x0 + x1) / 2, y0 + ph / 2,
                           f"{s['speed']} mph", size=9, bold=True,
                           color="#f59e0b", anchor="center")

    def _draw_func_class_bars(self, y0, mx, pw, sta_x, min_sta, max_sta, fcs, ph):
        for fc in fcs:
            x0, x1 = sta_x(fc["begin"]), sta_x(fc["end"])
            is_fwy = "freeway" in fc.get("class_type", "").lower()
            fill = "#dcfce7" if is_fwy else "#e0e7ff"
            border = "#10b981" if is_fwy else "#6366f1"
            class_type = fc.get("class_type", "")
            item = self._scene.addRect(
                QRectF(x0, y0 + 10, x1 - x0, ph - 20),
                QPen(QColor(border), 2), QBrush(QColor(fill)))
            self._set_hover(item,
                f"Functional Class: {class_type}"
                f"\nSta {_format_station(str(fc['begin']))} \u2013 "
                f"{_format_station(str(fc['end']))}")
            self._add_text((x0 + x1) / 2, y0 + ph / 2,
                           class_type, size=8, color=border,
                           anchor="center")

    # ── utilities ──────────────────────────────────────────────────────

    def _build_connections(self, d):
        conns = []
        for r in d.get("ramps", []):
            conns.append({
                "station": r["station"],
                "name": r.get("name", "Ramp"),
                "type": "ramp",
                "ramp_type": r.get("ramp_type", "entrance"),
            })
        conns.extend(d.get("intersections", []))
        return sorted(conns, key=lambda x: x["station"])

    def _set_hover(self, item, tooltip):
        """Enable hover tooltip and hand cursor on a graphics item."""
        item.setToolTip(tooltip)
        item.setAcceptHoverEvents(True)
        item.setCursor(Qt.PointingHandCursor)

    def _add_text(self, x, y, text, size=9, bold=False, color=TEXT_PRIMARY,
                  anchor="top-left"):
        item = QGraphicsSimpleTextItem(text)
        font = QFont("Segoe UI", size)
        if bold:
            font.setBold(True)
        item.setFont(font)
        item.setBrush(QColor(color))
        br = item.boundingRect()
        if "center" in anchor:
            x -= br.width() / 2
        elif "right" in anchor:
            x -= br.width()
        if "bottom" in anchor:
            y -= br.height()
        elif "center" == anchor:
            y -= br.height() / 2
        item.setPos(x, y)
        self._scene.addItem(item)
        return item
