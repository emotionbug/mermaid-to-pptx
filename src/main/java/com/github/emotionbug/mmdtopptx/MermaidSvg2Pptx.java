package com.github.emotionbug.mmdtopptx;

import io.github.bonigarcia.wdm.WebDriverManager;
import java.awt.*;
import java.awt.Shape;
import java.awt.font.FontRenderContext;
import java.awt.geom.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringWriter;
import java.nio.file.Files;
import java.util.*;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class MermaidSvg2Pptx {
  // ---------- Color / numeric parsing ----------
  static final Map<String, String> NAMED = new HashMap<>();
  private static final Logger log = LoggerFactory.getLogger(MermaidSvg2Pptx.class);

  static {
    // common Mermaid / CSS names
    NAMED.put("black", "000000");
    NAMED.put("white", "FFFFFF");
    NAMED.put("gray", "808080");
    NAMED.put("grey", "808080");
    NAMED.put("lightgrey", "D3D3D3");
    NAMED.put("lightgray", "D3D3D3");
    NAMED.put("darkgrey", "A9A9A9");
    NAMED.put("darkgray", "A9A9A9");
    NAMED.put("red", "FF0000");
    NAMED.put("green", "00FF00");
    NAMED.put("blue", "0000FF");
  }

  static ViewBox parseViewBox(String vb) {
    String[] p = vb.trim().split("[ ,]+");
    if (p.length != 4) throw new IllegalArgumentException("Invalid viewBox: " + vb);
    ViewBox v = new ViewBox();
    v.minX = Double.parseDouble(p[0]);
    v.minY = Double.parseDouble(p[1]);
    v.w = Double.parseDouble(p[2]);
    v.h = Double.parseDouble(p[3]);
    return v;
  }

  static String toHex(String c) {
    if (c == null) return null;
    c = c.trim();
    if (c.isEmpty() || isNoneOrTransparent(c)) return null;

    // Remove !important if present
    if (c.toLowerCase().contains("!important")) {
      c = c.substring(0, c.toLowerCase().indexOf("!important")).trim();
    }

    if (c.startsWith("#")) {
      String h = c.substring(1).toUpperCase();
      if (h.length() == 3) {
        return ""
            + h.charAt(0)
            + h.charAt(0)
            + h.charAt(1)
            + h.charAt(1)
            + h.charAt(2)
            + h.charAt(2);
      }
      return h;
    }
    String named = NAMED.get(c.toLowerCase());
    if (named != null) return named;

    // Try loose matching for rgb/rgba to handle various spacings or decimal points
    Matcher m =
        Pattern.compile("rgb\\(\\s*([\\d.]+)\\s*,\\s*([\\d.]+)\\s*,\\s*([\\d.]+)\\s*\\)")
            .matcher(c);
    if (m.matches()) {
      int r = (int) Math.round(Double.parseDouble(m.group(1)));
      int g = (int) Math.round(Double.parseDouble(m.group(2)));
      int b = (int) Math.round(Double.parseDouble(m.group(3)));
      return String.format("%02X%02X%02X", r, g, b);
    }
    Matcher ma =
        Pattern.compile(
                "rgba\\(\\s*([\\d.]+)\\s*,\\s*([\\d.]+)\\s*,\\s*([\\d.]+)\\s*,\\s*([\\d.]+)\\s*\\)")
            .matcher(c);
    if (ma.matches()) {
      int r = (int) Math.round(Double.parseDouble(ma.group(1)));
      int g = (int) Math.round(Double.parseDouble(ma.group(2)));
      int b = (int) Math.round(Double.parseDouble(ma.group(3)));
      double a = Double.parseDouble(ma.group(4));
      if (a == 0) return null; // Transparent
      return String.format("%02X%02X%02X", r, g, b);
    }
    Matcher hsm =
        Pattern.compile("hsl\\(\\s*([\\d.]+)\\s*,\\s*([\\d.]+)%\\s*,\\s*([\\d.]+)%\\s*\\)")
            .matcher(c);
    if (hsm.matches()) {
      double hh = Double.parseDouble(hsm.group(1));
      double ss = Double.parseDouble(hsm.group(2));
      double ll = Double.parseDouble(hsm.group(3));
      return hslToRgb(hh, ss, ll);
    }
    return null;
  }

  static String hslToRgb(double h, double s, double l) {
    s /= 100.0;
    l /= 100.0;
    double c = (1.0 - Math.abs(2.0 * l - 1.0)) * s;
    double x = c * (1.0 - Math.abs((h / 60.0) % 2.0 - 1.0));
    double m = l - c / 2.0;

    double r;
    double g;
    double b;

    if (h < 60) {
      r = c;
      g = x;
      b = 0;
    } else if (h < 120) {
      r = x;
      g = c;
      b = 0;
    } else if (h < 180) {
      r = 0;
      g = c;
      b = x;
    } else if (h < 240) {
      r = 0;
      g = x;
      b = c;
    } else if (h < 300) {
      r = x;
      g = 0;
      b = c;
    } else {
      r = c;
      g = 0;
      b = x;
    }
    return String.format(
        "%02X%02X%02X", (int) ((r + m) * 255), (int) ((g + m) * 255), (int) ((b + m) * 255));
  }

  static Double toPx(String s) {
    if (s == null) return null;
    s = s.trim();
    if (s.isEmpty()) return null;
    try {
      if (s.endsWith("px")) return Double.parseDouble(s.substring(0, s.length() - 2));
      if (s.endsWith("pt"))
        return Double.parseDouble(s.substring(0, s.length() - 2)) * (96.0 / 72.0);
      if (s.endsWith("em"))
        return Double.parseDouble(s.substring(0, s.length() - 2)); // caller interprets
      return Double.parseDouble(s);
    } catch (Exception ex) {
      return null;
    }
  }

  static StrokeStyle.LineDash dashFrom(String dashArray) {
    if (dashArray == null) return null;
    dashArray = dashArray.trim();
    if (dashArray.isEmpty() || isNoneOrTransparent(dashArray) || "0".equals(dashArray)) return null;
    String[] parts = dashArray.split("[ ,]+");
    double sum = 0;
    int cnt = 0;
    for (String p : parts) {
      p = p.trim();
      if (p.isEmpty()) continue;
      Double px = toPx(p);
      if (px != null) {
        sum += px;
        cnt++;
      }
    }
    if (cnt == 0) return null;
    double avg = sum / cnt;
    if (avg <= 2.2) return StrokeStyle.LineDash.DOT;
    return StrokeStyle.LineDash.DASH;
  }

  // ---------- SVG path parsing ----------
  static Shape parsePath(String d) {
    Path2D.Double path = new Path2D.Double();
    if (d == null || d.isEmpty()) return path;

    Matcher m =
        Pattern.compile("([a-zA-Z])|([-+]?(?:\\d*\\.\\d+|\\d+)(?:[eE][-+]?\\d+)?)").matcher(d);
    char cmd = ' ';
    List<Double> args = new ArrayList<>();
    double lastX = 0;
    double lastY = 0;

    double lastMoveX = 0;
    double lastMoveY = 0;

    while (m.find()) {
      if (m.group(1) != null) {
        cmd = m.group(1).charAt(0);
        args.clear();
        if (cmd == 'Z' || cmd == 'z') {
          path.closePath();
          lastX = lastMoveX;
          lastY = lastMoveY;
        }
      } else {
        args.add(Double.parseDouble(m.group(2)));
        if (isPathCmdReady(cmd, args)) {
          switch (cmd) {
            case 'M' -> {
              lastX = args.get(0);
              lastY = args.get(1);
              path.moveTo(lastX, lastY);
              lastMoveX = lastX;
              lastMoveY = lastY;
              cmd = 'L';
            }
            case 'm' -> {
              lastX += args.get(0);
              lastY += args.get(1);
              path.moveTo(lastX, lastY);
              lastMoveX = lastX;
              lastMoveY = lastY;
              cmd = 'l';
            }
            case 'L' -> {
              lastX = args.get(0);
              lastY = args.get(1);
              path.lineTo(lastX, lastY);
            }
            case 'l' -> {
              lastX += args.get(0);
              lastY += args.get(1);
              path.lineTo(lastX, lastY);
            }
            case 'H' -> {
              lastX = args.get(0);
              path.lineTo(lastX, lastY);
            }
            case 'h' -> {
              lastX += args.get(0);
              path.lineTo(lastX, lastY);
            }
            case 'V' -> {
              lastY = args.get(0);
              path.lineTo(lastX, lastY);
            }
            case 'v' -> {
              lastY += args.get(0);
              path.lineTo(lastX, lastY);
            }
            case 'C' -> {
              path.curveTo(
                  args.get(0), args.get(1), args.get(2), args.get(3), args.get(4), args.get(5));
              lastX = args.get(4);
              lastY = args.get(5);
            }
            case 'c' -> {
              path.curveTo(
                  lastX + args.get(0),
                  lastY + args.get(1),
                  lastX + args.get(2),
                  lastY + args.get(3),
                  lastX + args.get(4),
                  lastY + args.get(5));
              lastX += args.get(4);
              lastY += args.get(5);
            }
            case 'Q' -> {
              path.quadTo(args.get(0), args.get(1), args.get(2), args.get(3));
              lastX = args.get(2);
              lastY = args.get(3);
            }
            case 'q' -> {
              path.quadTo(
                  lastX + args.get(0),
                  lastY + args.get(1),
                  lastX + args.get(2),
                  lastY + args.get(3));
              lastX += args.get(2);
              lastY += args.get(3);
            }
            case 'A' -> {
              double rx = args.get(0), ry = args.get(1), rot = args.get(2);
              boolean largeArc = args.get(3) != 0, sweep = args.get(4) != 0;
              double nextX = args.get(5), nextY = args.get(6);
              svgArcTo(path, lastX, lastY, rx, ry, rot, largeArc, sweep, nextX, nextY);
              lastX = nextX;
              lastY = nextY;
            }
            case 'a' -> {
              double rx = args.get(0), ry = args.get(1), rot = args.get(2);
              boolean largeArc = args.get(3) != 0, sweep = args.get(4) != 0;
              double nextX = lastX + args.get(5), nextY = lastY + args.get(6);
              svgArcTo(path, lastX, lastY, rx, ry, rot, largeArc, sweep, nextX, nextY);
              lastX = nextX;
              lastY = nextY;
            }
          }
          args.clear();
        }
      }
    }
    return path;
  }

  private static void svgArcTo(
      Path2D.Double path,
      double x0,
      double y0,
      double rx,
      double ry,
      double xAxisRotation,
      boolean largeArcFlag,
      boolean sweepFlag,
      double x,
      double y) {
    if (x0 == x && y0 == y) return;
    if (rx == 0 || ry == 0) {
      path.lineTo(x, y);
      return;
    }
    rx = Math.abs(rx);
    ry = Math.abs(ry);

    double phi = Math.toRadians(xAxisRotation % 360);
    double cosPhi = Math.cos(phi);
    double sinPhi = Math.sin(phi);

    double dx2 = (x0 - x) / 2.0;
    double dy2 = (y0 - y) / 2.0;
    double x1p = cosPhi * dx2 + sinPhi * dy2;
    double y1p = -sinPhi * dx2 + cosPhi * dy2;

    double rx2 = rx * rx;
    double ry2 = ry * ry;
    double x1p2 = x1p * x1p;
    double y1p2 = y1p * y1p;

    double check = x1p2 / rx2 + y1p2 / ry2;
    if (check > 1.0) {
      rx *= Math.sqrt(check);
      ry *= Math.sqrt(check);
      rx2 = rx * rx;
      ry2 = ry * ry;
    }

    double sign = (largeArcFlag == sweepFlag) ? -1.0 : 1.0;
    double num = rx2 * ry2 - rx2 * y1p2 - ry2 * x1p2;
    double den = rx2 * y1p2 + ry2 * x1p2;
    double sq = Math.sqrt(Math.max(0, num / den));
    double cxp = sign * sq * (rx * y1p / ry);
    double cyp = sign * sq * (-ry * x1p / rx);

    double cx = cosPhi * cxp - sinPhi * cyp + (x0 + x) / 2.0;
    double cy = sinPhi * cxp + cosPhi * cyp + (y0 + y) / 2.0;

    double ux = (x1p - cxp) / rx;
    double uy = (y1p - cyp) / ry;
    double vx = (-x1p - cxp) / rx;
    double vy = (-y1p - cyp) / ry;

    double n = Math.sqrt(ux * ux + uy * uy);
    double p = ux;
    double angleStart = Math.toDegrees((uy < 0 ? -1.0 : 1.0) * Math.acos(p / n));

    n = Math.sqrt((ux * ux + uy * uy) * (vx * vx + vy * vy));
    p = ux * vx + uy * vy;
    double angleExtent =
        Math.toDegrees(
            (ux * vy - uy * vx < 0 ? -1.0 : 1.0) * Math.acos(Math.max(-1.0, Math.min(1.0, p / n))));

    if (!sweepFlag && angleExtent > 0) angleExtent -= 360;
    else if (sweepFlag && angleExtent < 0) angleExtent += 360;

    angleExtent %= 360;
    angleStart %= 360;

    Arc2D.Double arc =
        new Arc2D.Double(cx - rx, cy - ry, rx * 2, ry * 2, -angleStart, -angleExtent, Arc2D.OPEN);
    AffineTransform at = AffineTransform.getRotateInstance(phi, cx, cy);
    path.append(at.createTransformedShape(arc), true);
  }

  private static boolean isPathCmdReady(char cmd, List<Double> args) {
    return switch (Character.toUpperCase(cmd)) {
      case 'M', 'L' -> args.size() == 2;
      case 'H', 'V' -> args.size() == 1;
      case 'C' -> args.size() == 6;
      case 'Q' -> args.size() == 4;
      case 'A' -> args.size() == 7;
      default -> false;
    };
  }

  // ---------- Text measurement (approx) ----------
  static float stringWidthPx(String text, String fontName, float fontSizePx) {
    Font f = new Font(fontName, Font.PLAIN, Math.max(1, Math.round(fontSizePx)));
    FontRenderContext frc = new FontRenderContext(null, true, true);
    return (float) f.getStringBounds(text, frc).getWidth();
  }

  static String pickFontFamily(Map<String, String> computed) {
    String ff = computed.get("font-family");
    if (ff == null || ff.trim().isEmpty()) return "Malgun Gothic";
    // Use first family name
    String[] parts = ff.split(",");
    for (String p : parts) {
      p = p.replace("\"", "").trim();
      if (!p.isEmpty()) return p;
    }
    return "Malgun Gothic";
  }

  // ---------- Element drawing ----------
  static AffineTransform getElementTransform(Element el) {
    AffineTransform at = new AffineTransform();
    String transformStr = el.getAttribute("transform");
    if (transformStr.trim().isEmpty()) return at;

    Matcher m = Pattern.compile("(\\w+)\\s*\\(([^)]+)\\)").matcher(transformStr);
    while (m.find()) {
      String type = m.group(1).toLowerCase();
      String argsStr = m.group(2);
      String[] args = argsStr.split("[ ,]+");
      List<Double> d = new ArrayList<>();
      for (String arg : args) {
        arg = arg.trim();
        if (!arg.isEmpty()) d.add(Double.parseDouble(arg));
      }

      switch (type) {
        case "translate" -> {
          if (d.size() == 1) at.translate(d.get(0), 0);
          else if (d.size() >= 2) at.translate(d.get(0), d.get(1));
        }
        case "matrix" -> {
          if (d.size() >= 6)
            at.concatenate(
                new AffineTransform(d.get(0), d.get(1), d.get(2), d.get(3), d.get(4), d.get(5)));
        }
        case "scale" -> {
          if (d.size() == 1) at.scale(d.get(0), d.get(0));
          else if (d.size() >= 2) at.scale(d.get(0), d.get(1));
        }
        case "rotate" -> {
          if (d.size() == 1) at.rotate(Math.toRadians(d.get(0)));
          else if (d.size() >= 3) at.rotate(Math.toRadians(d.get(0)), d.get(1), d.get(2));
        }
        default -> log.debug("Skipping unsupported transform: {}", m.group(1));
      }
    }
    return at;
  }

  static AffineTransform getFullTransform(Element el, ComputedStyleResolver css) {
    if (css != null && css.hasBrowserStyles()) {
      Map<String, Object> data = css.getElementData(el);
      if (data != null && data.get("ctm") != null) {
        @SuppressWarnings("unchecked")
        Map<String, Object> c = (Map<String, Object>) data.get("ctm");
        return new AffineTransform(
            ((Number) c.get("a")).doubleValue(), ((Number) c.get("b")).doubleValue(),
            ((Number) c.get("c")).doubleValue(), ((Number) c.get("d")).doubleValue(),
            ((Number) c.get("e")).doubleValue(), ((Number) c.get("f")).doubleValue());
      }
    }

    AffineTransform at = new AffineTransform();
    List<Element> chain = new ArrayList<>();
    Node cur = el;
    while (cur instanceof Element e) {
      chain.add(e);
      cur = e.getParentNode();
    }
    Collections.reverse(chain);
    for (Element e : chain) {
      at.concatenate(getElementTransform(e));
    }
    return at;
  }

  static Point2D.Double getTransformedPoint(
      Element el, double x, double y, ComputedStyleResolver css) {
    AffineTransform at = getFullTransform(el, css);
    Point2D src = new Point2D.Double(x, y);
    Point2D dst = new Point2D.Double();
    at.transform(src, dst);
    return (Point2D.Double) dst;
  }

  static Rectangle2D getGlobalBBox(Element el, ComputedStyleResolver css) {
    Map<String, Object> data = css.getElementData(el);
    if (data == null || data.get("bbox") == null) return null;

    @SuppressWarnings("unchecked")
    Map<String, Object> b = (Map<String, Object>) data.get("bbox");
    double bx = ((Number) b.get("x")).doubleValue();
    double by = ((Number) b.get("y")).doubleValue();
    double bw = ((Number) b.get("width")).doubleValue();
    double bh = ((Number) b.get("height")).doubleValue();

    // line elements sometimes have 0 width/height in bbox but have coordinates
    if (bw == 0 && bh == 0 && !el.getTagName().equals("line")) return null;

    AffineTransform at = getFullTransform(el, css);
    Point2D[] pts =
        new Point2D[] {
          new Point2D.Double(bx, by),
          new Point2D.Double(bx + bw, by),
          new Point2D.Double(bx + bw, by + bh),
          new Point2D.Double(bx, by + bh)
        };

    double minX = Double.MAX_VALUE, minY = Double.MAX_VALUE;
    double maxX = -Double.MAX_VALUE, maxY = -Double.MAX_VALUE;
    for (Point2D p : pts) {
      at.transform(p, p);
      minX = Math.min(minX, p.getX());
      minY = Math.min(minY, p.getY());
      maxX = Math.max(maxX, p.getX());
      maxY = Math.max(maxY, p.getY());
    }
    return new Rectangle2D.Double(minX, minY, maxX - minX, maxY - minY);
  }

  static void drawRect(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      ComputedStyleResolver css) {
    double x0 = parseD(el.getAttribute("x"), 0);
    double y0 = parseD(el.getAttribute("y"), 0);
    double w0 = parseD(el.getAttribute("width"), 0);
    double h0 = parseD(el.getAttribute("height"), 0);
    double rx = parseD(el.getAttribute("rx"), 0);
    double ry = parseD(el.getAttribute("ry"), 0);

    Point2D.Double p0 = getTransformedPoint(el, x0, y0, css);
    Point2D.Double p1 = getTransformedPoint(el, x0 + w0, y0 + h0, css);

    double x = Math.min(p0.x, p1.x);
    double y = Math.min(p0.y, p1.y);
    double w = Math.abs(p1.x - p0.x);
    double h = Math.abs(p1.y - p0.y);

    XSLFAutoShape sh = container.createAutoShape();
    if (rx > 0 || ry > 0) {
      sh.setShapeType(ShapeType.ROUND_RECT);
    } else {
      sh.setShapeType(ShapeType.RECT);
    }
    sh.setAnchor(new Rectangle2D.Double(mp.x(x), mp.y(y), mp.w(w), mp.h(h)));
    applyShapeStyles(sh, mp, el, st);
  }

  static void drawCircle(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      ComputedStyleResolver css) {
    double cx = parseD(el.getAttribute("cx"), 0);
    double cy = parseD(el.getAttribute("cy"), 0);
    double r = parseD(el.getAttribute("r"), 0);

    Point2D.Double p0 = getTransformedPoint(el, cx - r, cy - r, css);
    Point2D.Double p1 = getTransformedPoint(el, cx + r, cy - r, css);
    Point2D.Double p2 = getTransformedPoint(el, cx + r, cy + r, css);
    Point2D.Double p3 = getTransformedPoint(el, cx - r, cy + r, css);

    double x = Math.min(Math.min(p0.x, p1.x), Math.min(p2.x, p3.x));
    double y = Math.min(Math.min(p0.y, p1.y), Math.min(p2.y, p3.y));
    double x_max = Math.max(Math.max(p0.x, p1.x), Math.max(p2.x, p3.x));
    double y_max = Math.max(Math.max(p0.y, p1.y), Math.max(p2.y, p3.y));

    double w = x_max - x;
    double h = y_max - y;

    XSLFAutoShape sh = container.createAutoShape();
    sh.setShapeType(ShapeType.ELLIPSE);
    sh.setAnchor(new Rectangle2D.Double(mp.x(x), mp.y(y), mp.w(w), mp.h(h)));
    applyShapeStyles(sh, mp, el, st);
  }

  static void drawEllipse(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      ComputedStyleResolver css) {
    double cx = parseD(el.getAttribute("cx"), 0);
    double cy = parseD(el.getAttribute("cy"), 0);
    double rx = parseD(el.getAttribute("rx"), 0);
    double ry = parseD(el.getAttribute("ry"), 0);

    Point2D.Double p0 = getTransformedPoint(el, cx - rx, cy - ry, css);
    Point2D.Double p1 = getTransformedPoint(el, cx + rx, cy - ry, css);
    Point2D.Double p2 = getTransformedPoint(el, cx + rx, cy + ry, css);
    Point2D.Double p3 = getTransformedPoint(el, cx - rx, cy + ry, css);

    double x = Math.min(Math.min(p0.x, p1.x), Math.min(p2.x, p3.x));
    double y = Math.min(Math.min(p0.y, p1.y), Math.min(p2.y, p3.y));
    double x_max = Math.max(Math.max(p0.x, p1.x), Math.max(p2.x, p3.x));
    double y_max = Math.max(Math.max(p0.y, p1.y), Math.max(p2.y, p3.y));

    double w = x_max - x;
    double h = y_max - y;

    XSLFAutoShape sh = container.createAutoShape();
    sh.setShapeType(ShapeType.ELLIPSE);
    sh.setAnchor(new Rectangle2D.Double(mp.x(x), mp.y(y), mp.w(w), mp.h(h)));
    applyShapeStyles(sh, mp, el, st);
  }

  static void applyShapeStyles(XSLFSimpleShape sh, Mapper mp, Element el, Map<String, String> st) {
    String fillStr = st.get("fill");
    String strokeStr = st.get("stroke");
    String swStr = st.get("stroke-width");

    String fill = toHex(fillStr);
    String stroke = toHex(strokeStr);
    Double sw = toPx(swStr);

    if (fill != null) {
      try {
        int rgb = Integer.parseInt(fill, 16);
        sh.setFillColor(new Color(rgb));
      } catch (Exception e) {
        // Ignore or log error
      }
    } else {
      // Default SVG fill is black if not specified.
      // But if explicitly "none", set to null (transparent).
      if (isNoneOrTransparent(fillStr)) {
        if (!(sh instanceof XSLFConnectorShape)) sh.setFillColor(null);
      } else if (fillStr == null) {
        // If not specified at all, SVG defaults to black.
        // However, for Mermaid, keeping it transparent (null) often matches expected output
        // if the background is handled elsewhere. Let's stay with null for now to avoid
        // everything turning black unless it's a known issue.
        if (!(sh instanceof XSLFConnectorShape)) sh.setFillColor(null);
      } else {
        // fillStr is not null but fill is null (parsing failed)
        // SVG default is black. Let's explicitly set it to black if it failed to parse but was
        // present.
        // UNLESS it's something like 'url(#...)' which we don't support yet.
        if (fillStr.contains("url(")) {
          if (!(sh instanceof XSLFConnectorShape)) sh.setFillColor(null);
        } else {
          if (!(sh instanceof XSLFConnectorShape)) sh.setFillColor(Color.BLACK);
        }
      }
    }

    if (stroke != null) {
      sh.setLineColor(new Color(Integer.parseInt(stroke, 16)));
      if (sw != null) sh.setLineWidth(Math.max(0.25, sw * mp.s));
      else sh.setLineWidth(mp.s);

      String dash = st.get("stroke-dasharray");
      StrokeStyle.LineDash ld = dashFrom(dash);
      if (ld != null) sh.setLineDash(ld);
    } else {
      sh.setLineWidth(0);
    }
  }

  static void drawLine(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      double x1,
      double y1,
      double x2,
      double y2,
      Map<String, Element> markers,
      ComputedStyleResolver css) {
    Point2D.Double p1 = getTransformedPoint(el, x1, y1, css);
    Point2D.Double p2 = getTransformedPoint(el, x2, y2, css);

    double sx1 = mp.x(p1.x), sy1 = mp.y(p1.y);
    double sx2 = mp.x(p2.x), sy2 = mp.y(p2.y);

    double x = Math.min(sx1, sx2);
    double y = Math.min(sy1, sy2);
    double w = Math.abs(sx2 - sx1);
    double h = Math.abs(sy2 - sy1);

    XSLFConnectorShape ln = container.createConnector();
    ln.setAnchor(new Rectangle2D.Double(x, y, w, h));

    if (sx1 > sx2) ln.setFlipHorizontal(true);
    if (sy1 > sy2) ln.setFlipVertical(true);

    applyShapeStyles(ln, mp, el, st);

    // markers
    String markerStart = st.get("marker-start");
    String markerEnd = st.get("marker-end");
    setMarkerDecoration(ln, markerStart, true, markers, css, el);
    setMarkerDecoration(ln, markerEnd, false, markers, css, el);

    drawOneMarker(container, mp, el, markerStart, x1, y1, markers, css);
    drawOneMarker(container, mp, el, markerEnd, x2, y2, markers, css);
  }

  static void setMarkerDecoration(
      XSLFSimpleShape sh,
      String url,
      boolean isStart,
      Map<String, Element> markers,
      ComputedStyleResolver css,
      Element hostEl) {
    if (url == null || !url.contains("#")) return;
    String id =
        url.substring(url.indexOf("#") + 1)
            .replace(")", "")
            .replace("'", "")
            .replace("\"", "")
            .trim();

    if (!id.toLowerCase().contains("arrowhead")
        && !id.toLowerCase().contains("filled-head")
        && !id.toLowerCase().contains("crosshead")) return;

    Element m = markers.get(id);
    if (m == null) {
      if (isStart) sh.setLineHeadDecoration(LineDecoration.DecorationShape.TRIANGLE);
      else sh.setLineTailDecoration(LineDecoration.DecorationShape.TRIANGLE);
      return;
    }

    // Determine shape
    LineDecoration.DecorationShape shape = LineDecoration.DecorationShape.TRIANGLE;
    if (id.toLowerCase().contains("crosshead")) {
      shape = LineDecoration.DecorationShape.ARROW;
    }

    if (isStart) sh.setLineHeadDecoration(shape);
    else sh.setLineTailDecoration(shape);

    // Determine size
    double mw = parseD(m.getAttribute("markerWidth"), 3);
    double mh = parseD(m.getAttribute("markerHeight"), 3);

    String mUnits = m.getAttribute("markerUnits");
    if (mUnits == null || mUnits.isEmpty()) mUnits = "strokeWidth";

    double sw = 1.0;
    String swStr = css.getStyle(hostEl, "stroke-width");
    Double dsw = toPx(swStr);
    if (dsw != null) sw = dsw;
    if (sw <= 0) sw = 1.0;

    double sizeValW = "userSpaceOnUse".equalsIgnoreCase(mUnits) ? mh / sw : mh;
    double sizeValL = "userSpaceOnUse".equalsIgnoreCase(mUnits) ? mw / sw : mw;

    LineDecoration.DecorationSize dsW = mapSize(sizeValW);
    LineDecoration.DecorationSize dsL = mapSize(sizeValL);

    if (isStart) {
      sh.setLineHeadWidth(dsW);
      sh.setLineHeadLength(dsL);
    } else {
      sh.setLineTailWidth(dsW);
      sh.setLineTailLength(dsL);
    }
  }

  static LineDecoration.DecorationSize mapSize(double val) {
    if (val <= 2.5) return LineDecoration.DecorationSize.SMALL;
    if (val <= 5.0) return LineDecoration.DecorationSize.MEDIUM;
    return LineDecoration.DecorationSize.LARGE;
  }

  static void drawOneMarker(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      String url,
      double x,
      double y,
      Map<String, Element> markers,
      ComputedStyleResolver css) {
    if (url == null || !url.contains("#")) return;
    String id =
        url.substring(url.indexOf("#") + 1)
            .replace(")", "")
            .replace("'", "")
            .replace("\"", "")
            .trim();

    // If it's an arrowhead, we handle it via POI LineDecoration instead of drawing it manually.
    // This avoids rotation issues with orient="auto" and ensures consistent arrow look.
    if (id.toLowerCase().contains("arrowhead")
        || id.toLowerCase().contains("filled-head")
        || id.toLowerCase().contains("crosshead")) return;

    Element m = markers.get(id);
    if (m == null) return;

    // markerUnits support
    String mUnits = m.getAttribute("markerUnits");
    if (mUnits == null || mUnits.isEmpty()) mUnits = "strokeWidth";
    double sw = 1.0;
    if ("strokeWidth".equalsIgnoreCase(mUnits)) {
      String swStr = css.getStyle(el, "stroke-width");
      Double dsw = toPx(swStr);
      if (dsw != null) sw = dsw;
      // Fallback for stroke-width="0" cases where marker is still intended to be visible
      if (sw <= 0) sw = 1.0;
    }

    double refX = parseD(m.getAttribute("refX"), 0) * sw;
    double refY = parseD(m.getAttribute("refY"), 0) * sw;
    Map<String, String> markerStyle = css.styleFor(m);

    NodeList children = m.getChildNodes();
    for (int i = 0; i < children.getLength(); i++) {
      Node n = children.item(i);
      if (!(n instanceof Element ce)) continue;

      Map<String, String> mst = new HashMap<>(markerStyle);
      mst.putAll(css.styleFor(ce));

      if (ce.getTagName().equals("circle")) {
        double cx = parseD(ce.getAttribute("cx"), 0) * sw;
        double cy = parseD(ce.getAttribute("cy"), 0) * sw;
        double r = parseD(ce.getAttribute("r"), 0) * sw;

        double dx = cx - refX;
        double dy = cy - refY;

        // Transform points of the circle's bounding box to handle scale/rotation of the host
        // element
        Point2D.Double p0 = getTransformedPoint(el, x + dx - r, y + dy - r, css);
        Point2D.Double p1 = getTransformedPoint(el, x + dx + r, y + dy - r, css);
        Point2D.Double p2 = getTransformedPoint(el, x + dx + r, y + dy + r, css);
        Point2D.Double p3 = getTransformedPoint(el, x + dx - r, y + dy + r, css);

        double tx = Math.min(Math.min(p0.x, p1.x), Math.min(p2.x, p3.x));
        double ty = Math.min(Math.min(p0.y, p1.y), Math.min(p2.y, p3.y));
        double tx_max = Math.max(Math.max(p0.x, p1.x), Math.max(p2.x, p3.x));
        double ty_max = Math.max(Math.max(p0.y, p1.y), Math.max(p2.y, p3.y));

        double tw = tx_max - tx;
        double th = ty_max - ty;

        XSLFAutoShape sh = container.createAutoShape();
        sh.setShapeType(ShapeType.ELLIPSE);
        sh.setAnchor(new Rectangle2D.Double(mp.x(tx), mp.y(ty), mp.w(tw), mp.h(th)));
        applyShapeStyles(sh, mp, ce, mst);
      } else if (ce.getTagName().equals("path") || ce.getTagName().equals("polygon")) {
        Shape mShape;
        if (ce.getTagName().equals("path")) {
          mShape = parsePath(ce.getAttribute("d"));
        } else {
          // polygon points
          Path2D.Double p2d = new Path2D.Double();
          String pts = ce.getAttribute("points");
          if (pts != null && !pts.trim().isEmpty()) {
            String[] pairs = pts.trim().split("\\s+");
            boolean first = true;
            for (String pair : pairs) {
              String[] xy = pair.split(",");
              if (xy.length == 2) {
                double px = Double.parseDouble(xy[0]);
                double py = Double.parseDouble(xy[1]);
                if (first) {
                  p2d.moveTo(px, py);
                  first = false;
                } else p2d.lineTo(px, py);
              }
            }
            p2d.closePath();
          }
          mShape = p2d;
        }

        // markerUnits scale
        if (sw != 1.0) {
          mShape = AffineTransform.getScaleInstance(sw, sw).createTransformedShape(mShape);
        }
        // Translate by -refX, -refY
        mShape = AffineTransform.getTranslateInstance(-refX, -refY).createTransformedShape(mShape);
        // Translate to host point (x, y)
        mShape = AffineTransform.getTranslateInstance(x, y).createTransformedShape(mShape);

        // Apply host element's full transform
        AffineTransform hostAt = getFullTransform(el, css);
        mShape = hostAt.createTransformedShape(mShape);

        XSLFFreeformShape free = container.createFreeform();
        Path2D.Double path = new Path2D.Double();
        PathIterator it = mShape.getPathIterator(null);
        double[] seg = new double[6];
        while (!it.isDone()) {
          int t = it.currentSegment(seg);
          if (t == PathIterator.SEG_MOVETO) path.moveTo(mp.x(seg[0]), mp.y(seg[1]));
          else if (t == PathIterator.SEG_LINETO) path.lineTo(mp.x(seg[0]), mp.y(seg[1]));
          else if (t == PathIterator.SEG_QUADTO)
            path.quadTo(mp.x(seg[0]), mp.y(seg[1]), mp.x(seg[2]), mp.y(seg[3]));
          else if (t == PathIterator.SEG_CUBICTO)
            path.curveTo(
                mp.x(seg[0]), mp.y(seg[1]), mp.x(seg[2]), mp.y(seg[3]), mp.x(seg[4]), mp.y(seg[5]));
          else if (t == PathIterator.SEG_CLOSE) path.closePath();
          it.next();
        }
        free.setPath(path);
        applyShapeStyles(free, mp, ce, mst);
      }
    }
  }

  static void drawPolygon(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      ComputedStyleResolver css) {
    String pts = el.getAttribute("points");
    if (pts == null || pts.trim().isEmpty()) return;
    String[] pairs = pts.trim().split("\\s+");

    Path2D.Double path = new Path2D.Double();
    AffineTransform at = getFullTransform(el, css);
    boolean first = true;

    for (String pair : pairs) {
      String[] xy = pair.split(",");
      if (xy.length != 2) continue;
      try {
        Point2D.Double p = new Point2D.Double(Double.parseDouble(xy[0]), Double.parseDouble(xy[1]));
        at.transform(p, p);
        if (first) {
          path.moveTo(mp.x(p.x), mp.y(p.y));
          first = false;
        } else {
          path.lineTo(mp.x(p.x), mp.y(p.y));
        }
      } catch (Exception ignore) {
      }
    }
    path.closePath();

    XSLFFreeformShape free = container.createFreeform();
    free.setPath(path);
    applyShapeStyles(free, mp, el, st);
  }

  static void drawPath(
      XSLFShapeContainer container,
      Mapper mp,
      Element el,
      Map<String, String> st,
      Map<String, Element> markers,
      ComputedStyleResolver css) {
    String d = el.getAttribute("d");
    if (d == null || d.trim().isEmpty()) return;

    Shape shape = parsePath(d);

    // Find start/end points for markers (untransformed)
    double[] firstPt = null;
    double[] lastPt = null;
    PathIterator pit = shape.getPathIterator(null);
    double[] coords = new double[6];
    while (!pit.isDone()) {
      int type = pit.currentSegment(coords);
      if (type == PathIterator.SEG_MOVETO || type == PathIterator.SEG_LINETO) {
        if (firstPt == null) firstPt = new double[] {coords[0], coords[1]};
        lastPt = new double[] {coords[0], coords[1]};
      } else if (type == PathIterator.SEG_QUADTO) {
        if (firstPt == null) firstPt = new double[] {coords[0], coords[1]};
        lastPt = new double[] {coords[2], coords[3]};
      } else if (type == PathIterator.SEG_CUBICTO) {
        if (firstPt == null) firstPt = new double[] {coords[0], coords[1]};
        lastPt = new double[] {coords[4], coords[5]};
      }
      pit.next();
    }

    AffineTransform at = getFullTransform(el, css);
    shape = at.createTransformedShape(shape);

    XSLFFreeformShape free = container.createFreeform();
    Path2D.Double path = new Path2D.Double();

    PathIterator it = shape.getPathIterator(null);
    double[] seg = new double[6];
    while (!it.isDone()) {
      int t = it.currentSegment(seg);
      if (t == PathIterator.SEG_MOVETO) path.moveTo(mp.x(seg[0]), mp.y(seg[1]));
      else if (t == PathIterator.SEG_LINETO) path.lineTo(mp.x(seg[0]), mp.y(seg[1]));
      else if (t == PathIterator.SEG_QUADTO)
        path.quadTo(mp.x(seg[0]), mp.y(seg[1]), mp.x(seg[2]), mp.y(seg[3]));
      else if (t == PathIterator.SEG_CUBICTO)
        path.curveTo(
            mp.x(seg[0]), mp.y(seg[1]), mp.x(seg[2]), mp.y(seg[3]), mp.x(seg[4]), mp.y(seg[5]));
      else if (t == PathIterator.SEG_CLOSE) path.closePath();
      it.next();
    }
    free.setPath(path);
    applyShapeStyles(free, mp, el, st);

    // markers
    String markerStart = st.get("marker-start");
    String markerEnd = st.get("marker-end");

    setMarkerDecoration(free, markerStart, true, markers, css, el);
    setMarkerDecoration(free, markerEnd, false, markers, css, el);

    if (firstPt != null)
      drawOneMarker(container, mp, el, markerStart, firstPt[0], firstPt[1], markers, css);
    if (lastPt != null)
      drawOneMarker(container, mp, el, markerEnd, lastPt[0], lastPt[1], markers, css);
  }

  static String extractText(Element textEl) {
    NodeList tspans = textEl.getElementsByTagName("tspan");
    if (tspans.getLength() > 0) {
      StringBuilder sb = new StringBuilder();
      for (int i = 0; i < tspans.getLength(); i++) {
        String t = tspans.item(i).getTextContent();
        if (t != null) {
          t = t.replace("\n", "").trim();
          if (!t.isEmpty()) {
            if (sb.length() > 0) sb.append("\n");
            sb.append(t);
          }
        }
      }
      return sb.toString();
    }
    String t = textEl.getTextContent();
    return t == null ? "" : t.trim();
  }

  static double computeDyPx(Element textEl, double fontSizePx) {
    double dy = 0;
    if (textEl.hasAttribute("dy")) {
      String v = textEl.getAttribute("dy").trim();
      if (v.endsWith("em")) dy += Double.parseDouble(v.substring(0, v.length() - 2)) * fontSizePx;
      else {
        Double px = toPx(v);
        if (px != null) dy += px;
      }
    }
    NodeList tspans = textEl.getElementsByTagName("tspan");
    if (tspans.getLength() > 0 && tspans.item(0) instanceof Element ts && ts.hasAttribute("dy")) {
      String v = ts.getAttribute("dy").trim();
      if (v.endsWith("em")) dy += Double.parseDouble(v.substring(0, v.length() - 2)) * fontSizePx;
      else {
        Double px = toPx(v);
        if (px != null) dy += px;
      }
    }
    return dy;
  }

  static void drawText(
      XSLFShapeContainer container,
      Mapper mp,
      List<Element> els,
      List<Map<String, String>> styles,
      ComputedStyleResolver css) {
    if (els.isEmpty()) return;

    Element el = els.get(0);
    Map<String, String> st = styles.get(0);

    StringBuilder sbText = new StringBuilder();
    for (Element e : els) {
      String t = extractText(e);
      if (!t.isEmpty()) {
        if (sbText.length() > 0) sbText.append("\n");
        sbText.append(t);
      }
    }
    String text = sbText.toString();
    if (text.isEmpty()) return;

    // font-size (px) -> pt
    Double fsPx = toPx(st.get("font-size"));
    if (fsPx == null) fsPx = 16.0;
    double fsPt = fsPx * mp.s * 0.5;

    String font = pickFontFamily(st);

    double minBoxX = Double.MAX_VALUE, minBoxY = Double.MAX_VALUE;
    double maxBoxX = -Double.MAX_VALUE, maxBoxY = -Double.MAX_VALUE;
    boolean hasBBox = false;

    TextParagraph.TextAlign align = TextParagraph.TextAlign.LEFT;

    for (Element e : els) {
      Map<String, Object> data = css.getElementData(e);
      if (data != null && data.get("bbox") != null) {
        @SuppressWarnings("unchecked")
        Map<String, Object> b = (Map<String, Object>) data.get("bbox");
        double bx = ((Number) b.get("x")).doubleValue();
        double by = ((Number) b.get("y")).doubleValue();
        double bw = ((Number) b.get("width")).doubleValue();
        double bh = ((Number) b.get("height")).doubleValue();

        AffineTransform at = getFullTransform(e, css);
        Rectangle2D localRect = new Rectangle2D.Double(bx, by, bw, bh);
        Shape transShape = at.createTransformedShape(localRect);
        Rectangle2D bounds = transShape.getBounds2D();

        minBoxX = Math.min(minBoxX, bounds.getX());
        minBoxY = Math.min(minBoxY, bounds.getY());
        maxBoxX = Math.max(maxBoxX, bounds.getMaxX());
        maxBoxY = Math.max(maxBoxY, bounds.getMaxY());
        hasBBox = true;

        String anchor = css.styleFor(e).get("text-anchor");
        if ("middle".equals(anchor)) align = TextParagraph.TextAlign.CENTER;
        else if ("end".equals(anchor)) align = TextParagraph.TextAlign.RIGHT;
      }
    }

    double boxX, boxY, boxW, boxH;
    if (hasBBox) {
      boxX = minBoxX;
      boxY = minBoxY;
      boxW = maxBoxX - minBoxX;
      boxH = maxBoxY - minBoxY;
    } else {
      // Fallback to manual measurement (simplified for group)
      double x = parseD(el.getAttribute("x"), -1e9);
      double y = parseD(el.getAttribute("y"), -1e9);

      if (x < -1e8 || y < -1e8) {
        NodeList ts = el.getElementsByTagName("tspan");
        for (int i = 0; i < ts.getLength(); i++) {
          Element t = (Element) ts.item(i);
          if (x < -1e8 && t.hasAttribute("x")) x = parseD(t.getAttribute("x"), 0);
          if (y < -1e8 && t.hasAttribute("y")) y = parseD(t.getAttribute("y"), 0);
        }
      }
      if (x < -1e8) x = 0;
      if (y < -1e8) y = 0;

      Point2D.Double pt = getTransformedPoint(el, x, y, css);
      x = pt.x;
      y = pt.y;
      y += computeDyPx(el, fsPx);

      String anchor = st.get("text-anchor");
      if (anchor == null || anchor.trim().isEmpty()) anchor = "start";
      anchor = anchor.trim().toLowerCase(Locale.ROOT);

      float wpx = stringWidthPx(text, font, fsPx.floatValue());
      boxW = wpx + 10;
      boxH = fsPx * 1.35 * (text.split("\n").length) + 6;

      if ("middle".equals(anchor)) {
        align = TextParagraph.TextAlign.CENTER;
        boxX = x - boxW / 2.0;
      } else if ("end".equals(anchor)) {
        align = TextParagraph.TextAlign.RIGHT;
        boxX = x - boxW;
      } else {
        align = TextParagraph.TextAlign.LEFT;
        boxX = x;
      }
      boxY = y - boxH / 2.0; // Center around y for fallback
    }

    XSLFTextBox tb = container.createTextBox();
    tb.setAnchor(new Rectangle2D.Double(mp.x(boxX), mp.y(boxY), mp.w(boxW), mp.h(boxH)));
    tb.setTextAutofit(XSLFTextShape.TextAutofit.NONE);
    tb.setWordWrap(false);
    tb.setInsets(new Insets2D(0, 0, 0, 0));
    tb.setVerticalAlignment(VerticalAlignment.MIDDLE);

    String[] lines = text.split("\n");
    for (int i = 0; i < lines.length; i++) {
      XSLFTextParagraph p = (i == 0) ? tb.getTextParagraphs().get(0) : tb.addNewTextParagraph();
      p.setTextAlign(align);
      XSLFTextRun r = p.addNewTextRun();
      r.setText(lines[i]);
      r.setFontFamily(font);
      r.setFontSize(fsPt);

      // Per-line color logic
      Element currentEl = (i < els.size()) ? els.get(i) : els.get(els.size() - 1);
      Map<String, String> currentSt =
          (i < styles.size()) ? styles.get(i) : styles.get(styles.size() - 1);

      String fillStr = currentSt.get("fill");
      String colorStr = currentSt.get("color");

      Element firstTspan = null;
      NodeList children = currentEl.getChildNodes();
      for (int j = 0; j < children.getLength(); j++) {
        Node n = children.item(j);
        if (n instanceof Element ce
            && (ce.getTagName().endsWith("tspan") || "tspan".equals(ce.getLocalName()))) {
          firstTspan = ce;
          break;
        }
      }

      if (firstTspan != null) {
        Map<String, String> tspanSt = css.styleFor(firstTspan);
        String tspanFill = tspanSt.get("fill");
        String tspanColor = tspanSt.get("color");
        if (tspanFill != null && !isNoneOrTransparent(tspanFill)) fillStr = tspanFill;
        if (tspanColor != null
            && !isNoneOrTransparent(tspanColor)
            && (colorStr == null || "rgb(0, 0, 0)".equals(colorStr))) colorStr = tspanColor;
      }

      String fill = toHex(fillStr);
      if ("ECECFF".equalsIgnoreCase(fill) || "EAEAEA".equalsIgnoreCase(fill)) {
        String c = toHex(colorStr);
        if ("000000".equalsIgnoreCase(c)) fill = c;
      }
      if (fill == null) fill = toHex(colorStr);
      if (fill != null) r.setFontColor(new Color(Integer.parseInt(fill, 16)));
    }
  }

  // ---------- Helpers ----------
  static double parseD(String s, double def) {
    Double d = parseD(s, (Double) def);
    return d == null ? def : d;
  }

  static Double parseD(String s, Double def) {
    if (s == null || s.trim().isEmpty()) return def;
    try {
      return Double.parseDouble(s.trim());
    } catch (Exception ex) {
      return def;
    }
  }

  // ---------- Main ----------
  private static String renderMermaidToSvg(String mmdContent) {
    WebDriverManager.edgedriver().setup();
    EdgeOptions options = new EdgeOptions();
    options.addArguments("--headless=new", "--disable-gpu", "--no-sandbox");
    WebDriver driver = new EdgeDriver(options);
    try {
      driver.get("about:blank");
      JavascriptExecutor js = (JavascriptExecutor) driver;

      js.executeScript(
          "var script = document.createElement('script');"
              + "script.src = 'https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js';"
              + "document.head.appendChild(script);");

      long start = System.currentTimeMillis();
      while (true) {
        Object defined = js.executeScript("return typeof mermaid !== 'undefined'");
        if (Boolean.TRUE.equals(defined)) break;
        if (System.currentTimeMillis() - start > 10000)
          throw new RuntimeException("Mermaid load timeout");
        Thread.sleep(100);
      }

      js.executeScript("mermaid.initialize({ startOnLoad: false });");

      Object result =
          js.executeAsyncScript(
              "var mmd = arguments[0];"
                  + "var callback = arguments[1];"
                  + "mermaid.render('mermaid-render-id', mmd).then(function(res) {"
                  + "  callback(res.svg);"
                  + "}).catch(function(err) {"
                  + "  callback('Error: ' + err.message);"
                  + "});",
              mmdContent);

      String svg = result.toString();
      if (svg.startsWith("Error:")) throw new RuntimeException("Mermaid rendering error: " + svg);
      return svg;
    } catch (Exception e) {
      throw new RuntimeException("Mermaid rendering failed", e);
    } finally {
      driver.quit();
    }
  }

  public static void main(String[] args) throws Exception {
    String inputPath = (args.length > 0) ? args[0] : "in.mmd";
    if (!new File(inputPath).exists() && args.length == 0) {
      inputPath = "in.svg";
    }

    String svgContent;
    if (inputPath.endsWith(".svg")) {
      svgContent = Files.readString(new File(inputPath).toPath());
    } else {
      log.info("Converting Mermaid MMD to SVG using Selenium...");
      String mmdContent = Files.readString(new File(inputPath).toPath());
      svgContent = renderMermaidToSvg(mmdContent);
      // Save to in.svg for debugging/persistence if needed, or just keep in memory
      Files.writeString(
          new File("tmp_" + System.currentTimeMillis() + ".svg").toPath(), svgContent);
    }

    String outPptx = "out_" + System.currentTimeMillis() + ".pptx";

    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
    dbf.setNamespaceAware(true);
    dbf.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
    dbf.setFeature("http://xml.org/sax/features/external-general-entities", false);
    dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false);

    Document doc =
        dbf.newDocumentBuilder()
            .parse(new java.io.ByteArrayInputStream(svgContent.getBytes("UTF-8")));

    Element svg = doc.getDocumentElement();
    ComputedStyleResolver css = new ComputedStyleResolver(doc);

    Map<String, Element> markers = new HashMap<>();
    NodeList mList = svg.getElementsByTagName("marker");
    for (int i = 0; i < mList.getLength(); i++) {
      Element m = (Element) mList.item(i);
      if (m.hasAttribute("id")) markers.put(m.getAttribute("id"), m);
    }

    // 1. Identify elements to draw
    List<Element> elementsToDraw;
    if (css.hasBrowserStyles()) {
      elementsToDraw = css.getOrderedElements();
    } else {
      elementsToDraw = new ArrayList<>();
      NodeList all = svg.getElementsByTagName("*");
      for (int i = 0; i < all.getLength(); i++) {
        Node n = all.item(i);
        if (!(n instanceof Element el)) continue;
        Node p = el.getParentNode();
        boolean skip = false;
        while (p != null) {
          if (p instanceof Element pe
              && (pe.getTagName().equals("defs") || pe.getTagName().equals("marker"))) {
            skip = true;
            break;
          }
          p = p.getParentNode();
        }
        if (!skip) elementsToDraw.add(el);
      }
    }

    // 2. Calculate bounding box of all drawn objects to determine slide size and centering
    double minX = Double.MAX_VALUE;
    double minY = Double.MAX_VALUE;
    double maxX = -Double.MAX_VALUE;
    double maxY = -Double.MAX_VALUE;
    boolean foundBox = false;

    for (Element el : elementsToDraw) {
      String tag = el.getTagName();
      if (tag.equals("rect")
          || tag.equals("circle")
          || tag.equals("ellipse")
          || tag.equals("polygon")
          || tag.equals("line")
          || tag.equals("path")
          || tag.equals("text")) {
        Rectangle2D b = getGlobalBBox(el, css);
        if (b != null) {
          minX = Math.min(minX, b.getMinX());
          minY = Math.min(minY, b.getMinY());
          maxX = Math.max(maxX, b.getMaxX());
          maxY = Math.max(maxY, b.getMaxY());
          foundBox = true;
        }
      }
    }

    ViewBox vb = new ViewBox();
    if (foundBox) {
      vb.minX = minX;
      vb.minY = minY;
      vb.w = maxX - minX;
      vb.h = maxY - minY;
    } else {
      String vbAttr = svg.getAttribute("viewBox");
      if (vbAttr != null && !vbAttr.isEmpty()) {
        ViewBox parsed = parseViewBox(vbAttr);
        vb.minX = parsed.minX;
        vb.minY = parsed.minY;
        vb.w = parsed.w;
        vb.h = parsed.h;
      } else {
        vb.minX = 0;
        vb.minY = 0;
        vb.w = 800;
        vb.h = 600;
      }
    }

    // 3. Determine slide size (Default 13.333x7.5 inches = 960x540 points, grow if needed)
    double margin = 40; // Total 40pt margin (20pt each side)
    double slideW = Math.max(13.333 * 72, vb.w + margin);
    double slideH = Math.max(7.5 * 72, vb.h + margin);

    log.debug("Content BBox: x={}, y={}, w={}, h={}", vb.minX, vb.minY, vb.w, vb.h);
    log.debug("Setting slide size to: {}x{}", slideW, slideH);

    try (XMLSlideShow ppt = new XMLSlideShow()) {
      ppt.setPageSize(new Dimension((int) Math.ceil(slideW), (int) Math.ceil(slideH)));
      XSLFSlide slide = ppt.createSlide();

      Mapper mp = new Mapper(vb, slideW, slideH);

      drawRecursive(slide, svg, mp, css, markers);

      try (FileOutputStream fos = new FileOutputStream(outPptx)) {
        ppt.write(fos);
      }
    }
    log.info("Wrote to {}", outPptx);
  }

  private static void drawRecursive(
      XSLFShapeContainer container,
      Element parent,
      Mapper mp,
      ComputedStyleResolver css,
      Map<String, Element> markers) {
    NodeList nl = parent.getChildNodes();
    for (int i = 0; i < nl.getLength(); i++) {
      Node node = nl.item(i);
      if (!(node instanceof Element el)) continue;

      String tag = el.getTagName();
      if (tag.equals("defs")
          || tag.equals("marker")
          || tag.equals("style")
          || tag.equals("metadata")) continue;

      Map<String, String> st = css.styleFor(el);

      switch (tag) {
        case "rect" -> drawRect(container, mp, el, st, css);
        case "circle" -> drawCircle(container, mp, el, st, css);
        case "ellipse" -> drawEllipse(container, mp, el, st, css);
        case "polygon" -> drawPolygon(container, mp, el, st, css);
        case "line" -> {
          double x1 = parseD(el.getAttribute("x1"), 0);
          double y1 = parseD(el.getAttribute("y1"), 0);
          double x2 = parseD(el.getAttribute("x2"), 0);
          double y2 = parseD(el.getAttribute("y2"), 0);
          drawLine(container, mp, el, st, x1, y1, x2, y2, markers, css);
        }
        case "path" -> drawPath(container, mp, el, st, markers, css);
        case "text" -> {
          List<Element> textGroup = new ArrayList<>();
          List<Map<String, String>> styleGroup = new ArrayList<>();
          textGroup.add(el);
          styleGroup.add(st);

          // Try to find adjacent text elements to group
          int j = i + 1;
          while (j < nl.getLength()) {
            Node nextNode = nl.item(j);
            if (!(nextNode instanceof Element nextEl)) {
              j++;
              continue;
            }
            if (!nextEl.getTagName().equals("text")) break;

            Map<String, String> nextSt = css.styleFor(nextEl);

            // Check grouping criteria
            if (canGroupText(el, st, nextEl, nextSt, css)) {
              textGroup.add(nextEl);
              styleGroup.add(nextSt);
              j++;
            } else {
              break;
            }
          }
          if (textGroup.size() > 1) {
            log.info(
                "Grouping {} text elements starting with: '{}'", textGroup.size(), extractText(el));
          }
          drawText(container, mp, textGroup, styleGroup, css);
          i = j - 1; // Skip grouped elements
        }
        default -> drawRecursive(container, el, mp, css, markers);
      }
    }
  }

  private static boolean canGroupText(
      Element el1,
      Map<String, String> st1,
      Element el2,
      Map<String, String> st2,
      ComputedStyleResolver css) {
    // 1. Check style similarity
    if (!Objects.equals(st1.get("font-size"), st2.get("font-size"))) return false;
    if (!Objects.equals(st1.get("font-weight"), st2.get("font-weight"))) return false;
    if (!Objects.equals(st1.get("text-anchor"), st2.get("text-anchor"))) return false;
    if (!Objects.equals(st1.get("font-family"), st2.get("font-family"))) return false;

    // 2. Check center point (x coordinate) similarity
    Double x1 = getX(el1);
    Double x2 = getX(el2);
    if (x1 == null || x2 == null) return false;
    if (Math.abs(x1 - x2) > 2.0) return false;

    // 3. Check adjacency (y coordinate)
    Double y1 = getY(el1);
    Double y2 = getY(el2);
    if (y1 == null || y2 == null) return false;

    Double fsPx = toPx(st1.get("font-size"));
    if (fsPx == null) fsPx = 16.0;

    double diffY = Math.abs(y1 - y2);
    // y 좌표 차이가 폰트 크기의 0.5배 ~ 3배 사이여야 인접한 줄로 판단
    if (diffY < fsPx * 0.5 || diffY > fsPx * 3.0) return false;

    return true;
  }

  private static Double getX(Element el) {
    String xAttr = el.getAttribute("x");
    if (xAttr == null || xAttr.isEmpty()) {
      NodeList ts = el.getElementsByTagName("tspan");
      if (ts.getLength() > 0) xAttr = ((Element) ts.item(0)).getAttribute("x");
    }
    return parseD(xAttr, null);
  }

  private static Double getY(Element el) {
    String yAttr = el.getAttribute("y");
    if (yAttr == null || yAttr.isEmpty()) {
      NodeList ts = el.getElementsByTagName("tspan");
      if (ts.getLength() > 0) yAttr = ((Element) ts.item(0)).getAttribute("y");
    }
    return parseD(yAttr, null);
  }

  // ---------- CSS extraction using Batik Bridge and Headless Browser ----------
  static void instrumentSvg(
      Element el, AtomicInteger counter, Map<String, Element> indexToElement) {
    String idx = String.valueOf(counter.getAndIncrement());
    el.setAttribute("data-style-idx", idx);
    indexToElement.put(idx, el);
    NodeList nl = el.getChildNodes();
    for (int i = 0; i < nl.getLength(); i++) {
      if (nl.item(i) instanceof Element ce) {
        instrumentSvg(ce, counter, indexToElement);
      }
    }
  }

  static String documentToString(Document doc) throws TransformerException {
    TransformerFactory tf = TransformerFactory.newInstance();
    Transformer transformer = tf.newTransformer();
    StringWriter writer = new StringWriter();
    transformer.transform(new DOMSource(doc), new StreamResult(writer));
    return writer.getBuffer().toString();
  }

  private static boolean isNoneOrTransparent(String color) {
    return "none".equalsIgnoreCase(color) || "transparent".equalsIgnoreCase(color);
  }

  // ---------- Geometry mapping: SVG viewBox units -> PPT inches (WIDE 13.333 x 7.5) ----------
  static class ViewBox {
    double minX, minY, w, h;
  }

  static class Mapper {
    final ViewBox vb;
    final double s, ox, oy;

    Mapper(ViewBox vb, double slideW, double slideH) {
      this.vb = vb;
      double sx = (vb.w > 0) ? slideW / vb.w : 1.0;
      double sy = (vb.h > 0) ? slideH / vb.h : 1.0;
      // Maintain 1:1 scale if it fits or we grew the slide, scale down only if absolutely
      // necessary.
      // With the new main logic, sx/sy will be >= 1.0 always when growing is enabled.
      this.s = Math.min(1.0, Math.min(sx, sy));
      this.ox = (slideW - vb.w * s) / 2.0;
      this.oy = (slideH - vb.h * s) / 2.0;
    }

    double x(double v) {
      return (v - vb.minX) * s + ox;
    }

    double y(double v) {
      return (v - vb.minY) * s + oy;
    }

    double w(double v) {
      return v * s;
    }

    double h(double v) {
      return v * s;
    }
  }

  static class ComputedStyleResolver {
    private final Map<String, Element> indexToElement = new HashMap<>();
    private final Map<String, Map<String, Object>> indexToData = new HashMap<>();
    private Map<String, Map<String, String>> browserStyles = null;
    private List<Map<String, Object>> elementDataList = null;

    public ComputedStyleResolver(Document doc) {
      // 1. Instrument SVG with Indices for mapping
      instrumentSvg(doc.getDocumentElement(), new AtomicInteger(0), indexToElement);

      // 2. Selenium extraction
      try {
        extractStylesWithSelenium(doc);
        if (elementDataList != null) {
          log.info("Selenium data extracted successfully. Count: {}", elementDataList.size());
        }
      } catch (Exception t) {
        log.error("Selenium extraction failed: {}", t.getMessage(), t);
      }
    }

    @SuppressWarnings("unchecked")
    private void extractStylesWithSelenium(Document doc) throws TransformerException, IOException {
      WebDriverManager.edgedriver().setup();

      EdgeOptions options = new EdgeOptions();
      options.addArguments(
          "--headless=new", // SVG / CSS 계산 정확도 ↑
          "--disable-gpu",
          "--window-size=1920,1080",
          "--disable-dev-shm-usage",
          "--no-sandbox");

      WebDriver driver = new EdgeDriver(options);
      try {
        String xml = documentToString(doc);
        File tempFile = File.createTempFile("mermaid-instr-", ".svg");
        try {
          Files.writeString(tempFile.toPath(), xml);
          driver.get(tempFile.toURI().toString());

          JavascriptExecutor js = (JavascriptExecutor) driver;
          String script =
              "const results = [];"
                  + "const all = document.querySelectorAll('*');"
                  + "all.forEach(el => {"
                  + "    const idx = el.getAttribute('data-style-idx');"
                  + "    if (idx === null) return;"
                  + "    let p = el.parentElement;"
                  + "    let inDefs = false;"
                  + "    while (p) {"
                  + "        const tn = p.tagName.toLowerCase();"
                  + "        if (tn === 'defs' || tn === 'marker') { inDefs = true; break; }"
                  + "        p = p.parentElement;"
                  + "    }"
                  + "    const s = window.getComputedStyle(el);"
                  + "    const styles = {"
                  + "        'fill': s.fill,"
                  + "        'stroke': s.stroke,"
                  + "        'stroke-width': s.strokeWidth,"
                  + "        'stroke-dasharray': s.strokeDasharray,"
                  + "        'font-size': s.fontSize,"
                  + "        'font-family': s.fontFamily,"
                  + "        'text-anchor': s.textAnchor,"
                  + "        'marker-start': s.markerStart,"
                  + "        'marker-end': s.markerEnd,"
                  + "        'color': s.color"
                  + "    };"
                  + "    const attrs = {};"
                  + "    for (let i = 0; i < el.attributes.length; i++) {"
                  + "        attrs[el.attributes[i].name] = el.attributes[i].value;"
                  + "    }"
                  + "    let bbox = null;"
                  + "    let ctm = null;"
                  + "    if (typeof el.getBBox === 'function') {"
                  + "        try {"
                  + "            const b = el.getBBox();"
                  + "            bbox = { x: b.x, y: b.y, width: b.width, height: b.height };"
                  + "            const c = el.getCTM();"
                  + "            if (c) ctm = { a: c.a, b: c.b, c: c.c, d: c.d, e: c.e, f: c.f };"
                  + "        } catch(e) {}"
                  + "    }"
                  + "    results.push({ idx, tagName: el.tagName.toLowerCase(), styles, attrs, bbox, ctm, isHidden: inDefs });"
                  + "});"
                  + "return results;";

          this.elementDataList = (List<Map<String, Object>>) js.executeScript(script);
          this.browserStyles = new HashMap<>();
          for (Map<String, Object> item : elementDataList) {
            String idx = (String) item.get("idx");
            this.browserStyles.put(idx, (Map<String, String>) item.get("styles"));
            this.indexToData.put(idx, item);
          }
        } finally {
          tempFile.delete();
        }
      } finally {
        driver.quit();
      }
    }

    public boolean hasBrowserStyles() {
      return elementDataList != null;
    }

    public List<Element> getOrderedElements() {
      if (elementDataList == null) return null;
      List<Element> result = new ArrayList<>();
      for (Map<String, Object> item : elementDataList) {
        if (Boolean.TRUE.equals(item.get("isHidden"))) continue;
        String idx = (String) item.get("idx");
        Element el = indexToElement.get(idx);
        if (el != null) result.add(el);
      }
      return result;
    }

    public Map<String, Object> getElementData(Element el) {
      return indexToData.get(el.getAttribute("data-style-idx"));
    }

    public String getStyle(Element el, String propertyName) {
      String val = null;

      // 1. Try Browser Styles first (via index)
      if (browserStyles != null) {
        String idx = el.getAttribute("data-style-idx");
        if (browserStyles.containsKey(idx)) {
          val = browserStyles.get(idx).get(propertyName);
          if (isNoneOrTransparent(val) || "rgba(0, 0, 0, 0)".equalsIgnoreCase(val)) val = null;
        }
      }

      // 2. Fallback to Attribute
      if (val == null || isNoneOrTransparent(val)) {
        String attr = el.getAttribute(propertyName);
        if (!attr.isEmpty()) val = attr;
      }
      return val;
    }

    public Map<String, String> styleFor(Element el) {
      Map<String, String> merged = new HashMap<>();
      String[] pAttrs = {
        "fill",
        "stroke",
        "stroke-width",
        "stroke-dasharray",
        "font-size",
        "font-family",
        "text-anchor",
        "marker-start",
        "marker-end",
        "color"
      };
      for (String a : pAttrs) {
        String val = getStyle(el, a);
        if (val != null) merged.put(a, val);
      }
      return merged;
    }
  }
}
