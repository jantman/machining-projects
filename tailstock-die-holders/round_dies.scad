// Large round fitting in a #5 holder 1.522" diameter and 0.485" deep
// Small round fitting in a #3 holder 1.020" diameter and 0.365" deep

$fn = 720;

stock_dia = 2.000;
length = 2.000;
small_bore_dia = 1.020;
small_depth = 0.365;
large_bore_dia = 1.522;
large_depth = 0.485;
padding = stock_dia - large_bore_dia;
full_dia_length = 1.25;
bore_dia = 0.501;
tommy_dia = 0.251;
tommy_backset = length / 2;
set_screw_hole_dia = 0.1610; // 10-24

difference() {
    // body
    union() {
        cylinder(d=stock_dia, h=full_dia_length);
        translate([0, 0, full_dia_length]) {
            cylinder(d1=stock_dia, d2=(small_bore_dia + padding), h=(length - full_dia_length));
        }
    }
    // center bore
    translate([0, 0, -0.1]) {
        cylinder(d=bore_dia, h=length + 0.2);
    }
    // large recess
    translate([0, 0, -0.1]) {
        cylinder(d=large_bore_dia, h=large_depth + 0.1);
    }
    // small recess
    translate([0, 0, length - small_depth]) {
        cylinder(d=small_bore_dia, h=small_depth + 0.1);
    }
    // tommy bar bore
    translate([0,((stock_dia / 2) + 0.1), tommy_backset]) {
        rotate([90, 0, 0]) {
            cylinder(d=tommy_dia, h=stock_dia + 0.2);
        }
    }
    // 180-degree set screw holes on large bore
    translate([0,((stock_dia / 2) + 0.1), large_depth / 2]) {
        rotate([90, 0, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
    // 90-degree set screw hole on large bore
    translate([0, 0, large_depth / 2]) {
        rotate([0, 90, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
    // 180-degree set screw holes on small bore
    translate([0,((stock_dia / 2) + 0.1), length - (small_depth / 2)]) {
        rotate([90, 0, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
    // 90-degree set screw hole on small bore
    translate([0, 0, length - (small_depth / 2)]) {
        rotate([0, 90, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
}
