package ir.dotin.utils.xls.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Enumeration;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

/**
 * list resources available from the classpath @ *
 */
public class ResourceUtil {

    /**
     * for all elements of java.class.path get a Collection of resources Pattern
     * pattern = Pattern.compile(".*"); gets all resources
     *
     * @param pattern the pattern to match
     * @return the resources in the order they are found
     */
    public static Collection<InputStream> getResources(final Pattern pattern) throws IOException {
        final ArrayList<InputStream> retval = new ArrayList<InputStream>();
        final String classPath = System.getProperty("java.class.path", ".");
        final String[] classPathElements = classPath.split(System.getProperty("path.separator"));
        for (final String element : classPathElements) {
            retval.addAll(getResources(element, pattern));
        }
        return retval;
    }

    public static Collection<InputStream> getResourcesWithBaseClassPath(String baseClassPath, final Pattern pattern) throws IOException {
        final ArrayList<InputStream> retval = new ArrayList<InputStream>();
        retval.addAll(getResources(baseClassPath, pattern));
        return retval;
    }

    private static Collection<InputStream> getResources(final String element, final Pattern pattern) throws IOException {
        final ArrayList<InputStream> retval = new ArrayList<InputStream>();
        final File file = new File(element);
        if (file.isDirectory()) {
            retval.addAll(getResourcesFromDirectory(file, pattern));
        } else {
            retval.addAll(getResourcesFromJarFile(file, pattern));
        }
        return retval;
    }

    private static Collection<InputStream> getResourcesFromJarFile(
            final File file,
            final Pattern pattern) throws IOException {
        final ArrayList<InputStream> retval = new ArrayList<InputStream>();
        ZipFile zf;
        try {
            zf = new ZipFile(file);
        } catch (final ZipException e) {
            throw new Error(e);
        } catch (final IOException e) {
            throw new Error(e);
        }
        final Enumeration e = zf.entries();
        boolean closeZF = true;
        while (e.hasMoreElements()) {
            final ZipEntry ze = (ZipEntry) e.nextElement();
            final String fileName = ze.getName();
            final boolean accept = pattern.matcher(fileName).matches();
            if (accept) {
                retval.add(zf.getInputStream(ze));
                closeZF= false;
            }
        }
        try {
            if (closeZF) {
                zf.close();
            }
        } catch (final IOException e1) {
            throw new Error(e1);
        }
        return retval;
    }

    private static Collection<InputStream> getResourcesFromDirectory(
            final File directory,
            final Pattern pattern) {
        final Collection<InputStream> retval = new ArrayList<InputStream>();
        final File[] fileList = directory.listFiles();
        for (final File file : fileList) {
            if (file.isDirectory()) {
                retval.addAll(getResourcesFromDirectory(file, pattern));
            } else {
                try {
                    final String fileName = file.getCanonicalPath();
                    if (fileName.endsWith(".jar")){
                        retval.addAll(getResourcesFromJarFile(file,pattern));
                    }else {
                        final boolean accept = pattern.matcher(fileName).matches();
                        if (accept) {
                            retval.add(new FileInputStream(fileName));
                        }
                    }
                } catch (final IOException e) {
                    throw new Error(e);
                }
            }
        }
        return retval;
    }


}
