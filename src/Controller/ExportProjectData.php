<?php

/*
 * This file is part of the Kimai time-tracking app.
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace App\Controller;

use App\Project\ProjectService;
use App\Project\ProjectStatisticService;
use App\Reporting\ProjectDetails\ProjectDetailsQuery;
use App\Repository\Query\TimesheetQuery;
use App\Repository\TimesheetRepository;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

/**
 * Controller used to render reports.
 */
final class ExportProjectData extends AbstractController
{
    public function __construct(protected TimesheetRepository $repository, protected ProjectService $projectService, protected ProjectStatisticService $serviceStat)
    {
    }

    #[Route(path: '/admin/export_project_data/{project_name}', name: 'export_project_data', methods: ['GET'])]
    public function defaultReport(string $project_name): Response
    {
        $dateFactory = $this->getDateTimeFactory();
        $user = $this->getUser();
        $project = $this->projectService->findProjectByName($project_name);
        $query = new ProjectDetailsQuery($dateFactory->createDateTime(), $user);
        $query->setProject($project);
        $projectDetails = $this->serviceStat->getProjectsDetails($query);
        $usersStats = $projectDetails->getUserStats();
        $activities = $projectDetails->getActivities();

        $querySheet = new TimesheetQuery();
        if (!empty($project)) {
            $querySheet->setProjects([$project]);
        } else {
            return new Response('No data');
        }
        $querySheet->setName('MyTimesListing');

        $result = $this->repository->getTimesheetResult($querySheet)->getResults();
        if (!empty($result)) {
            if (count($result) === 0) {
                return new Response('No data');
            }
        } else {
            return new Response('No data');
        }

        $inputFileName = $this->getParameter('kernel.project_dir') . '/public/template_report_project.xlsx';
        $spreadsheet = IOFactory::load($inputFileName);
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('B2', 'SUIVI DES CHARGES PROJET ' . strtoupper($project->getName()));
        $row = 7;
        $nbActivities = count($activities);
        $compteurActivities = 1;
        foreach ($activities as $activity) {
            $sheet->setCellValue("B{$row}", $activity->getName());
            $sheet->setCellValue("G{$row}", round($activity->getActivity()->getTimeBudget() / 3600, 2));
            $sheet->setCellValue("H{$row}", 0);
            if ($compteurActivities < $nbActivities) {
                $row++;
                $compteurActivities++;
                $sheet->insertNewRowBefore($row);
                $sheet->mergeCells("B{$row}:F{$row}");
            }
        }
        $finalRowCount = $row + 1;
        $sheet->setCellValue("G{$finalRowCount}", "=SUM(G7:G{$row})");
        $sheet->setCellValue("H{$finalRowCount}", "=SUM(H7:H{$row})");

        $newest = $result[0]->getEnd();
        $oldest = $result[0]->getBegin();
        foreach ($result as $obj) {
            if ($obj->getEnd() > $newest) {
                $newest = $obj->getEnd();
            }
            if ($obj->getBegin() < $oldest) {
                $oldest = $obj->getBegin();
            }
        }

        $startDate = clone $oldest;
        $endDate = clone $newest;
        $startDate = $startDate->format('Y-m-d');
        $numeroSemaine = date('W', strtotime($startDate));
        $diff = strtotime($startDate) - strtotime($endDate->format('Y-m-d'));
        $nombreDeSemainesDiff = abs(floor($diff / (7 * 24 * 60 * 60)));

        $col = 'I';
        $compteurCol = 1;
        $nbCol = $nombreDeSemainesDiff + 1;

        for ($i = 0; $i <= $nombreDeSemainesDiff; $i++) {
            $currentSemaine = ((int)$numeroSemaine + $i);
            $currentSheet = array_filter($result, function ($value) use ($currentSemaine) {
                return (int)date('W', strtotime($value->getBegin()->format('Y-m-d'))) === $currentSemaine;
            });
            $sheet->setCellValue("{$col}5", 'W' . $currentSemaine);
            $user_col = $col;
            $compteurUser = 1;
            $nbUsers = count($usersStats);

            foreach ($usersStats as $userStat) {
                $sheet->setCellValue("{$user_col}6", $userStat->getUser()->getName());
                $sheet->setCellValue("{$user_col}{$finalRowCount}", "=SUM({$user_col}7:{$user_col}{$row})");
                $sheet->getStyle("{$user_col}{$finalRowCount}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $currentUserSheet = array_filter($currentSheet, function ($value) use ($userStat) {
                    return $userStat->getUser()->getId() === $value->getUser()->getId();
                });

                $actCount = 0;
                for ($y = 7; $y <= $row; $y++) {
                    $activitiesSheet = array_filter($currentUserSheet, function ($value) use ($activities, $actCount) {
                        return $value->getActivity()->getId() === $activities[$actCount]->getActivity()->getId();
                    });
                    $countDuration = array_reduce($activitiesSheet, function ($carry, $element) {
                        return $carry + $element->getDuration();
                    }, 0);
                    $countDuration = round($countDuration / 3600, 2);
                    $sheet->setCellValue("{$user_col}{$y}", "{$countDuration}");
                    $actCount++;
                }

                if ($compteurUser < $nbUsers) {
                    $user_col++;
                    $compteurUser++;
                    $sheet->insertNewColumnBefore($user_col);
                }
            }
            $sheet->mergeCells("{$col}5:{$user_col}5");
            $sheet->getStyle("{$col}5")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $col = $user_col;
            if ($compteurCol < $nbCol) {
                $compteurCol++;
                $col++;
                $sheet->insertNewColumnBefore($col);
            }

        }

        $styleArrayThin = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];

        $styleArrayLarge = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];

        $EndCol1 = $col;
        $sheet->getStyle("I5:{$EndCol1}5")->applyFromArray($styleArrayLarge);
        $col++;
        for ($x = 7; $x <= $row; $x++) {
            $sheet->setCellValue("{$col}{$x}", "=SUM(I{$x}:{$EndCol1}{$x})");
        }
        $sheet->setCellValue("{$col}{$finalRowCount}", "=SUM({$col}7:{$col}{$row})");
        $EndCol1++;
        $col++;
        for ($x = 7; $x <= $row; $x++) {
            $sheet->setCellValue("{$col}{$x}", "=SUM({$EndCol1}{$x} + H{$x}) - G{$x}");
        }
        $sheet->setCellValue("{$col}{$finalRowCount}", "=SUM({$col}7:{$col}{$row})");
        $startColForStyle = 'A';
        while (true) {
            $sheet->getColumnDimension($startColForStyle)->setAutoSize(true);
            if ($startColForStyle === $col) {
                break;
            }
            $startColForStyle++;
        }
        $sheet->getStyle("B6:{$col}6")->applyFromArray($styleArrayLarge);
        $sheet->getStyle("B7:{$col}{$row}")->applyFromArray($styleArrayThin);
        $row++;
        $sheet->getStyle("C{$row}:{$col}{$row}")->applyFromArray($styleArrayLarge);

        $writer = new Xlsx($spreadsheet);
        $filename = "export_project_{$project_name}.xlsx";
        $filepath = "/tmp/{$filename}";
        $writer->save($filepath);

        $response = new Response();
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response->headers->set('Content-Disposition', "attachment;filename={$filename}");
        $response->headers->set('Cache-Control', 'max-age=0');
        $response->headers->set('Content-length', filesize($filepath));
        $response->sendHeaders();
        $response->setContent(file_get_contents($filepath));
        unlink($filepath);

        return $response;
    }
}
