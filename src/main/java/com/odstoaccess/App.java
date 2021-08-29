package com.odstoaccess;

public final class App {
    public static void main(String[] args) throws Exception {
        ODSToAccess reader;

        if (args.length == 3) {
            try {
                reader = new ODSToAccess(args[0], args[1], args[2]);
                reader.readAllLines();
            } catch (Exception e) {
                System.out.println("Program couldn't execute properly : " + e.getMessage());

                return;
            }
        }

        else {
            throw new IllegalArgumentException(
                    "Must have 3 arguments, the ODS file, then the MSAccess file, then the database name");
        }
    }
}