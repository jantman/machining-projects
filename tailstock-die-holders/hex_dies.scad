// Large hex fitting in a hex holder 1.425" across flats (1.645 across points) and 0.520" deep
// Small hex fitting in a hex holder 1.011" across flats (1.167 across points) and 0.355" deep

$fn = 720;

stock_dia = 2.000;
length = 2.000;
small_bore_dia = 1.167; // 1.011 across flats is 1.167 across points
small_depth = 0.355;
large_bore_dia = 1.645; // 1.425 across flats is 1.645 across points
large_depth = 0.520;
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
        cylinder(d=large_bore_dia, h=large_depth + 0.1, $fn=6);
    }
    // small recess
    translate([0, 0, length - small_depth]) {
        cylinder(d=small_bore_dia, h=small_depth + 0.1, $fn=6);
    }
    // tommy bar bore
    translate([0,((stock_dia / 2) + 0.1), tommy_backset]) {
        rotate([90, 0, 0]) {
            cylinder(d=tommy_dia, h=stock_dia + 0.2);
        }
    }
    // 180-degree set screw holes on large bore
    translate([-1 * ((stock_dia / 2) + 0.1),0, large_depth / 2]) {
        rotate([00, 90, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
    // 180-degree set screw holes on small bore
    translate([-1 * ((stock_dia / 2) + 0.1),0, length - (small_depth / 2)]) {
        rotate([00, 90, 0]) {
            cylinder(d=set_screw_hole_dia, h=stock_dia + 0.2);
        }
    }
}
